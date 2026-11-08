"""回归测试：Esc 键有时无法停止回放

问题背景：
- 原实现用 kb.add_hotkey('esc', ...) 注册 Esc。
  keyboard 库的单键热键只有在"当前按下的按键集合恰好只有该键"时才会触发，
  只要同时按住 Shift / Ctrl / Alt 等任意其它按键，热键就会被**静默丢弃**。
  实测：仅按 Esc 会触发，Shift+Esc、Ctrl+Esc、A+Esc 均不触发。
  于是"按住 Shift 再按 Esc"这类操作根本停不回放。

修复（两条互相备份的路径）：
1. 改用 kb.hook() 原始事件钩子逐个按键事件判断 Esc，不受其它按键影响；
2. 回放线程在 _should_abort_playback 中直接查询系统物理按键状态
   （GetAsyncKeyState），即使 keyboard 的全局钩子失效也能中止回放。

覆盖：
1. 不再使用 add_hotkey('esc', ...) 这种单键热键
2. 原始钩子在"同时按住其它键"时依然能识别 Esc
3. Esc 键按下事件能触发停止回放；非 Esc 键、抬起事件不触发
4. 回放线程能通过物理按键状态检测中止回放，并记录中止原因
"""
import os
import sys
import time
from threading import Lock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import auto_test_gui as ag

from test_playback_stop import (  # 复用回放测试替身
    Fake,
    FakeBtn,
    FakeList,
    FakePyAutoGUI,
    FakeRoot,
    FakeVar,
    make_loop_group,
    wait_until,
)


# --------------------------------------------------------------------------
# 替身
# --------------------------------------------------------------------------
class FakeEvent:
    """模拟 keyboard 的按键事件"""

    def __init__(self, event_type, name):
        self.event_type = event_type
        self.name = name


def make_key_event(name, event_type):
    """构造 keyboard 库真实格式的按键事件（用于走完整的热键匹配逻辑）"""
    from keyboard._keyboard_event import KeyboardEvent

    scan_codes = {'esc': 1, 'shift': 42, 'ctrl': 29, 'a': 30}
    return KeyboardEvent(
        event_type=event_type,
        scan_code=scan_codes.get(name, 0),
        name=name,
        is_keypad=False,
    )


class EscFake(Fake):
    """在回放测试替身基础上，记录 Esc 相关调用"""

    def __init__(self):
        super().__init__()
        self.esc_stop_calls = 0
        self.escape_pressed = False

    def stop_by_esc(self):
        self.esc_stop_calls += 1
        super().stop_by_esc()

    def _is_escape_pressed(self):
        """替换真实实现，模拟 Esc 键被按下的物理状态"""
        return self.escape_pressed


# --------------------------------------------------------------------------
# 用例
# --------------------------------------------------------------------------
def test_esc_not_registered_as_single_key_hotkey():
    """不得再使用 kb.add_hotkey('esc', ...) 注册 Esc（该方式会被静默丢弃）"""
    app = EscFake()
    registered = []
    hooked = []
    old_add_hotkey = ag.kb.add_hotkey
    old_hook = ag.kb.hook
    try:
        ag.kb.add_hotkey = lambda hotkey, *a, **k: registered.append(hotkey)
        ag.kb.hook = lambda callback, *a, **k: hooked.append(callback) or (lambda: None)

        app._register_hotkeys()

        assert 'esc' not in registered, f"仍用单键热键注册了 Esc: {registered}"
        assert 'ctrl+shift+l' in registered, f"录制快捷键丢失: {registered}"
        assert 'ctrl+shift+p' in registered, f"回放快捷键丢失: {registered}"
        assert hooked, "未注册原始按键事件钩子"
        assert app._on_global_key_event in hooked, "钩子未指向 Esc 处理函数"
    finally:
        ag.kb.add_hotkey = old_add_hotkey
        ag.kb.hook = old_hook
    print("[OK] Esc 改用原始事件钩子，不再使用会被丢弃的单键热键")


def test_raw_hook_detects_esc_while_other_key_held():
    """核心回归：同时按住 Shift 时，Esc 依然必须被识别

    直接走 keyboard 库真实的热键匹配逻辑（direct_callback），
    验证我们的原始钩子能收到事件 —— 这正是 add_hotkey('esc') 会丢失的组合。
    """
    app = EscFake()
    old_add_hotkey = ag.kb.add_hotkey
    old_hook = ag.kb.hook
    try:
        ag.kb.add_hotkey = lambda hotkey, *a, **k: (lambda: None)
        # 使用真实的 kb.hook，让事件真正进入 keyboard 的分发流程
        ag.kb.hook = old_hook
        app._register_hotkeys()
        time.sleep(0.3)   # 等待 keyboard 的处理线程就绪

        listener = ag.kb._listener
        listener.start_if_necessary()   # 确保监听器已初始化
        from keyboard._keyboard_event import KEY_DOWN, KEY_UP

        # 先按住 Shift，再按 Esc —— add_hotkey('esc') 在此组合下会被丢弃
        listener.direct_callback(make_key_event('shift', KEY_DOWN))
        listener.direct_callback(make_key_event('esc', KEY_DOWN))
        time.sleep(0.4)
        listener.direct_callback(make_key_event('esc', KEY_UP))
        listener.direct_callback(make_key_event('shift', KEY_UP))
        time.sleep(0.2)

        assert app.esc_stop_calls == 1, \
            f"按住 Shift 时 Esc 未被识别（停止调用 {app.esc_stop_calls} 次）"
    finally:
        ag.kb.add_hotkey = old_add_hotkey
        ag.kb.hook = old_hook
    print("[OK] 同时按住其它键时，Esc 依然能被识别并停止回放")


def test_esc_event_dispatch_rules():
    """只有 Esc 的"按下"事件才触发停止；其它按键与抬起事件不触发"""
    app = EscFake()

    app._on_global_key_event(FakeEvent('down', 'esc'))
    assert app.esc_stop_calls == 1, "Esc 按下未触发停止"

    app._on_global_key_event(FakeEvent('up', 'esc'))
    app._on_global_key_event(FakeEvent('down', 'a'))
    app._on_global_key_event(FakeEvent('down', 'enter'))
    assert app.esc_stop_calls == 1, \
        f"非 Esc 按键或抬起事件误触发停止: {app.esc_stop_calls}"
    print("[OK] 仅 Esc 的按下事件触发停止")


def test_escape_key_state_polled_by_playback():
    """回放线程必须能通过物理按键状态检测 Esc 并中止回放"""
    gui = FakePyAutoGUI()
    old_pyautogui = ag.pyautogui
    ag.pyautogui = gui
    try:
        app = EscFake()
        app.actions = [make_loop_group(loop_count=200, delay=0.05)]

        app._start_playback_after_recording_stop()
        thread = app.playback_thread
        assert wait_until(lambda: gui.clicks > 0), "回放未开始"

        app.escape_pressed = True      # 模拟用户按下 Esc
        thread.join(timeout=3.0)

        assert not thread.is_alive(), "回放线程未退出"
        assert app.is_playing is False, "停止后播放标志仍为 True"
        assert app._playback_abort_reason and "Esc" in app._playback_abort_reason, \
            f"未记录 Esc 中止原因: {app._playback_abort_reason}"

        clicks_after_stop = gui.clicks
        time.sleep(0.3)
        assert gui.clicks == clicks_after_stop, "中止后仍执行了动作"

        app.root.flush()
        assert app.play_btn.cget("text") == "开始回放", "按钮未复位"
    finally:
        ag.pyautogui = old_pyautogui
    print("[OK] 回放线程可通过物理按键状态检测 Esc 并立即中止")


def test_escape_state_query_does_not_raise():
    """_is_escape_pressed 必须在没有真实按键时安全返回布尔值"""
    app = EscFake()
    app._is_escape_pressed = ag.AutoTestGUI._is_escape_pressed.__get__(app)
    assert app._is_escape_pressed() in (True, False), "返回值不是布尔值"
    print("[OK] 物理按键状态查询安全可用")


def run():
    test_esc_not_registered_as_single_key_hotkey()
    test_raw_hook_detects_esc_while_other_key_held()
    test_esc_event_dispatch_rules()
    test_escape_key_state_polled_by_playback()
    test_escape_state_query_does_not_raise()
    print("\n全部通过")


if __name__ == "__main__":
    run()
