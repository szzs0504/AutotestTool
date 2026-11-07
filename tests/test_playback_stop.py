"""回归测试：循环体（loop_group）中 Esc / 鼠标左上角故障保护无法停止回放

问题背景：
- 循环组分支使用不可中断的 time.sleep()，且完全没有鼠标左上角检测；
- stop_playback() 只置 is_playing=False，不递增播放代号，
  旧线程在新一次回放开始后会"复活"继续执行，并覆盖新回放的按钮状态。

覆盖：
1. stop_playback() 递增播放代号，使正在运行的线程立即失效
2. 循环体中按 Esc（stop_by_esc）能停止回放，且延时可被立即打断
3. 循环体中把鼠标移到左上角能停止回放
4. 失效的旧线程不能覆盖新一次回放的按钮状态与播放标志
5. 长延时可被立即打断（停止耗时远小于延时本身）
6. 没有可回放动作时，按钮不会卡在"停止回放"
"""
import os
import sys
import time
from threading import Lock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import auto_test_gui as ag


# --------------------------------------------------------------------------
# 测试替身
# --------------------------------------------------------------------------
class FakeRoot:
    """模拟 Tk root.after：把回调排队，由测试显式 flush（可控执行顺序）"""

    def __init__(self):
        self.queue = []

    def after(self, delay_ms, func):
        self.queue.append(func)
        return len(self.queue)

    def flush(self):
        pending, self.queue = self.queue, []
        for func in pending:
            func()


class FakeBtn:
    def __init__(self, text="", state="normal"):
        self._text = text
        self._state = state

    def config(self, **kw):
        if "text" in kw:
            self._text = kw["text"]
        if "state" in kw:
            self._state = kw["state"]

    def cget(self, key):
        return {"text": self._text, "state": self._state}[key]


class FakeVar:
    def __init__(self, value=None):
        self._value = value

    def get(self):
        return self._value

    def set(self, value):
        self._value = value


class FakeList:
    def __init__(self):
        self.items = []

    def delete(self, *a):
        self.items = []

    def insert(self, *a):
        self.items.append(a)

    def see(self, *a):
        pass


class FakePyAutoGUI:
    """记录鼠标/键盘调用，并允许模拟鼠标位置（左上角故障保护）"""

    def __init__(self):
        self.FAILSAFE = True
        self._pos = (500, 500)
        self.clicks = 0
        self.moves = 0
        self.writes = 0

    # -- 位置 --
    def position(self):
        return self._pos

    def move_to_corner(self):
        self._pos = (0, 0)

    # -- 动作 --
    def moveTo(self, x, y, duration=0):
        self.moves += 1

    def click(self, *a, **k):
        self.clicks += 1

    def doubleClick(self, *a, **k):
        self.clicks += 1

    def scroll(self, *a, **k):
        pass

    def write(self, *a, **k):
        self.writes += 1

    def hotkey(self, *a, **k):
        pass

    def press(self, *a, **k):
        pass


class Fake(ag.AutoTestGUI):
    """跳过真实 __init__（需要管理员权限与完整 GUI），只构造回放所需字段"""

    def __init__(self):
        self.root = FakeRoot()
        self.actions_lock = Lock()
        self.test_index_lock = Lock()
        self._variable_counter_lock = Lock()
        self.test_index = 0
        self.is_playing = False
        self.is_recording = False
        self._playback_abort_reason = None
        self.actions = []
        self.ui_elements = []
        self.play_btn = FakeBtn(text="开始回放")
        self.record_btn = FakeBtn(text="开始录制")
        self.status_var = FakeVar("")
        self.playback_list = FakeList()
        self.loop_count = FakeVar(1)
        self.loop_interval = FakeVar(0.0)
        self.interval_random = FakeVar(False)
        self.interval_min = FakeVar(0.5)
        self.interval_max = FakeVar(1.5)


def make_loop_group(loop_count=50, delay=0.05):
    return {
        "type": "loop_group",
        "loop_count": loop_count,
        "loop_interval_type": "fixed",
        "loop_interval": {"value": 0.0},
        "loop_actions": [
            {"type": "click", "x": 100, "y": 100, "button": "left", "delay": delay},
        ],
    }


def wait_until(predicate, timeout=8.0, interval=0.02):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if predicate():
            return True
        time.sleep(interval)
    return False


# --------------------------------------------------------------------------
# 用例
# --------------------------------------------------------------------------
def test_stop_bumps_generation():
    """stop_playback() 必须递增播放代号，使旧线程立即失效"""
    app = Fake()
    app.is_playing = True
    app.test_index = 3
    stale_index = 3

    app.stop_playback()

    assert app.test_index == 4, f"播放代号未递增: {app.test_index}"
    assert app.is_playing is False, "停止后播放标志仍为 True"
    # 旧线程应立即判定自己失效
    assert app._is_playback_valid(stale_index) is False, "旧线程未失效"
    assert app._should_abort_playback(stale_index) is True, "旧线程未判定为应中止"
    print("[OK] stop_playback() 递增播放代号，旧线程立即失效")


def test_stale_worker_cannot_clobber_new_playback():
    """失效的旧线程收尾时，不得覆盖新一次回放的按钮状态与播放标志"""
    app = Fake()
    app.is_playing = True
    app.test_index = 7
    app.play_btn = FakeBtn(text="停止回放")

    # 旧线程（代号 5）尝试收尾：应被完全忽略
    app._end_playback(5, "回放已中止")
    assert app.is_playing is True, "旧线程不应复位播放标志"
    assert app.play_btn.cget("text") == "停止回放", "旧线程不应覆盖新回放的按钮"
    assert app.root.queue == [], "旧线程不应排队任何 UI 更新"

    # 代号 7 的收尾回调排队后，被代号 8 的新回放取代：回调执行时必须自行校验
    app._end_playback(7, "执行完成")
    assert len(app.root.queue) == 1, "当前代号的收尾应排队一个 UI 回调"
    app.test_index = 8
    app.root.flush()
    assert app.play_btn.cget("text") == "停止回放", "过期回调覆盖了新回放的按钮"

    # 当前代号的收尾回调必须生效
    app._end_playback(8, "执行完成")
    app.root.flush()
    assert app.play_btn.cget("text") == "开始回放", "当前代号的收尾未生效"
    print("[OK] 失效的旧线程无法覆盖新回放的按钮状态与播放标志")


def test_esc_stops_loop_group():
    """循环体中按 Esc（stop_by_esc）必须停止回放，且延时立即被打断

    关键断言：Esc 之后一个动作都不允许再执行。
    旧实现的循环体分支使用不可中断的 time.sleep()，
    醒来后会无条件再执行一个动作 —— 这里 clicks 必须恒为 0。
    """
    gui = FakePyAutoGUI()
    old_pyautogui = ag.pyautogui
    ag.pyautogui = gui
    try:
        app = Fake()
        # 循环体内每个动作延时 1.0 秒：足以让线程停在延时里，
        # 保证 Esc 落在延时期间（此时还没有任何点击）
        app.actions = [make_loop_group(loop_count=200, delay=1.0)]

        app._start_playback_after_recording_stop()
        thread = app.playback_thread

        # 回放线程先有 0.5 秒的窗口切换等待，再进入循环体的 1 秒延时，
        # 因此 1.0 秒时线程正停在循环体内的延时里
        time.sleep(1.0)
        assert thread.is_alive(), "回放线程意外退出"
        assert gui.clicks == 0, "延时期间不应有点击"

        t0 = time.monotonic()
        app.stop_by_esc()          # 与键盘热键完全相同的路径
        elapsed = time.monotonic() - t0

        thread.join(timeout=3.0)
        assert not thread.is_alive(), "回放线程未退出"
        assert elapsed < 1.0, f"停止耗时过长（延时未被立即打断）: {elapsed:.2f}s"
        assert app.is_playing is False, "停止后播放标志仍为 True"

        # 核心断言：Esc 之后一个动作都不允许再执行
        assert gui.clicks == 0, f"Esc 之后仍执行了 {gui.clicks} 个动作"

        # 停止后不应再有新的动作
        clicks_after_stop = gui.clicks
        time.sleep(0.3)
        assert gui.clicks == clicks_after_stop, \
            f"停止后仍执行了动作: {clicks_after_stop} -> {gui.clicks}"

        app.root.flush()
        assert app.play_btn.cget("text") == "开始回放", "按钮未复位"
    finally:
        ag.pyautogui = old_pyautogui
    print("[OK] 循环体中按 Esc 可立即停止回放（延时被打断，无后续动作）")


def test_failsafe_stops_loop_group():
    """循环体中把鼠标移到屏幕左上角必须停止回放"""
    gui = FakePyAutoGUI()
    old_pyautogui = ag.pyautogui
    ag.pyautogui = gui
    try:
        app = Fake()
        app.actions = [make_loop_group(loop_count=200, delay=0.05)]

        app._start_playback_after_recording_stop()
        thread = app.playback_thread

        assert wait_until(lambda: gui.clicks > 0), "循环体未开始执行"

        gui.move_to_corner()      # 鼠标移到左上角
        thread.join(timeout=3.0)

        assert not thread.is_alive(), "回放线程未退出"
        assert app.is_playing is False, "停止后播放标志仍为 True"
        assert app._playback_abort_reason is not None, "未记录中止原因"

        clicks_after_stop = gui.clicks
        time.sleep(0.3)
        assert gui.clicks == clicks_after_stop, \
            f"故障保护触发后仍执行了动作: {clicks_after_stop} -> {gui.clicks}"

        app.root.flush()
        assert app.play_btn.cget("text") == "开始回放", "按钮未复位"
    finally:
        ag.pyautogui = old_pyautogui
    print("[OK] 循环体中鼠标移到左上角可停止回放")


def test_long_delay_is_interrupted():
    """延时动作必须可中断：停止耗时应远小于延时本身"""
    gui = FakePyAutoGUI()
    old_pyautogui = ag.pyautogui
    ag.pyautogui = gui
    try:
        app = Fake()
        # 循环体内部每个动作延时 5 秒
        app.actions = [make_loop_group(loop_count=200, delay=5.0)]

        app._start_playback_after_recording_stop()
        thread = app.playback_thread

        assert wait_until(lambda: app.is_playing), "回放未启动"
        time.sleep(1.0)  # 让线程越过启动等待，进入 5 秒延时

        t0 = time.monotonic()
        app.stop_playback()
        elapsed = time.monotonic() - t0
        thread.join(timeout=3.0)

        assert not thread.is_alive(), "回放线程未退出"
        assert elapsed < 1.0, f"5 秒延时未被立即打断，停止耗时 {elapsed:.2f}s"
        assert gui.clicks == 0, "延时被打断后仍执行了动作"
    finally:
        ag.pyautogui = old_pyautogui
    print("[OK] 循环体内的长延时可被立即打断")


def test_empty_actions_resets_ui():
    """没有可回放动作时，按钮不能卡在"停止回放" """
    app = Fake()
    app.actions = []
    app._start_playback_after_recording_stop()
    thread = app.playback_thread
    thread.join(timeout=3.0)
    assert not thread.is_alive(), "回放线程未退出"
    app.root.flush()
    assert app.play_btn.cget("text") == "开始回放", "按钮未复位"
    assert app.is_playing is False, "播放标志未复位"
    print("[OK] 空动作列表时按钮与播放标志正确复位")


def test_esc_then_new_playback_keeps_button_stopped():
    """用户报告的核心症状：按 Esc 后立即开始新回放，按钮不能变成"开始回放"

    旧播放线程在停止后仍存活片刻，若它的收尾逻辑写入 UI，
    会把新一次回放的"停止回放"按钮覆盖成"开始回放"，
    用户看到的现象就是"回放过程中按钮有时是开始回放"。
    """
    gui = FakePyAutoGUI()
    old_pyautogui = ag.pyautogui
    ag.pyautogui = gui
    try:
        app = Fake()
        app.actions = [make_loop_group(loop_count=200, delay=0.05)]

        # 第一次回放
        app._start_playback_after_recording_stop()
        first_thread = app.playback_thread
        assert wait_until(lambda: gui.clicks > 0), "第一次回放未开始"

        # 按 Esc 停止（不等旧线程退出）
        app.stop_by_esc()

        # 立刻开始第二次回放 —— 旧线程此时很可能仍在运行
        app._start_playback_after_recording_stop()
        second_thread = app.playback_thread
        assert second_thread is not first_thread, "第二次回放未创建新线程"

        # 让旧线程彻底退出，并给它足够时间去写 UI
        first_thread.join(timeout=3.0)
        assert not first_thread.is_alive(), "旧回放线程未退出"
        time.sleep(0.3)

        # 第二次回放期间，按钮必须始终显示"停止回放"
        app.root.flush()
        assert app.play_btn.cget("text") == "停止回放", \
            f"旧线程把新回放的按钮覆盖成了 {app.play_btn.cget('text')!r}"
        assert app.is_playing is True, "播放标志被旧线程复位"

        # 第二次回放也应能被 Esc 正常停止
        app.stop_by_esc()
        second_thread.join(timeout=3.0)
        assert not second_thread.is_alive(), "第二次回放线程未退出"
        app.root.flush()
        assert app.play_btn.cget("text") == "开始回放", "停止后按钮未复位"
    finally:
        ag.pyautogui = old_pyautogui
    print("[OK] Esc 后立即开始新回放，旧线程不会覆盖新回放的按钮状态")


def test_failsafe_during_delay_in_loop_group():
    """循环体延时进行中把鼠标移到左上角，必须立即中止"""
    gui = FakePyAutoGUI()
    old_pyautogui = ag.pyautogui
    ag.pyautogui = gui
    try:
        app = Fake()
        # 循环体内每个动作延时 3 秒
        app.actions = [make_loop_group(loop_count=200, delay=3.0)]

        app._start_playback_after_recording_stop()
        thread = app.playback_thread

        # 等待线程进入延时（此时还没有任何点击）
        time.sleep(0.8)
        assert gui.clicks == 0, "延时被打断前不应有点击"

        gui.move_to_corner()          # 延时进行中把鼠标移到左上角
        thread.join(timeout=3.0)

        assert not thread.is_alive(), "故障保护未在延时中生效"
        assert app.is_playing is False, "停止后播放标志仍为 True"
        assert gui.clicks == 0, "故障保护后仍执行了动作"

        app.root.flush()
        assert app.play_btn.cget("text") == "开始回放", "按钮未复位"
    finally:
        ag.pyautogui = old_pyautogui
    print("[OK] 循环体延时进行中鼠标移到左上角可立即中止")


def test_failsafe_flag_restored_after_stop():
    """停止后必须把 pyautogui 的全局故障保护开关恢复回来

    回放开始时会在锁内关闭 FAILSAFE（否则鼠标移到左上角会抛异常），
    回放结束后必须重新开启，否则故障保护会永久失效。
    """
    gui = FakePyAutoGUI()
    old_pyautogui = ag.pyautogui
    ag.pyautogui = gui
    try:
        app = Fake()
        app.actions = [make_loop_group(loop_count=200, delay=0.05)]

        app._start_playback_after_recording_stop()
        thread = app.playback_thread
        assert wait_until(lambda: gui.clicks > 0), "回放未开始"
        assert gui.FAILSAFE is False, "回放期间故障保护应被关闭"

        app.stop_by_esc()
        thread.join(timeout=3.0)
        assert not thread.is_alive(), "回放线程未退出"

        # 停止后开关必须恢复（无论走的是哪条收尾路径）
        assert gui.FAILSAFE is True, "停止后故障保护开关未恢复"
    finally:
        ag.pyautogui = old_pyautogui
    print("[OK] 停止回放后 pyautogui 故障保护开关已恢复")


def run():
    test_stop_bumps_generation()
    test_stale_worker_cannot_clobber_new_playback()
    test_esc_stops_loop_group()
    test_failsafe_stops_loop_group()
    test_long_delay_is_interrupted()
    test_empty_actions_resets_ui()
    test_esc_then_new_playback_keeps_button_stopped()
    test_failsafe_during_delay_in_loop_group()
    test_failsafe_flag_restored_after_stop()
    print("\n全部通过")


if __name__ == "__main__":
    run()
