"""回归测试："插入动作"对话框（单击 / 双击 / 右键 / 回车）

覆盖：
1. 四个选项各自构造出正确的动作字典（build_insert_action）
2. 对话框中确实存在四个选项，且默认选中"单击"
3. 选项行位于"点击位置"行的上方
4. 选中不同选项后，确定按钮插入的动作与选项一致
5. 回车动作回放时先点击目标位置再按下回车键
6. 右键动作回放时使用右键；双击动作回放时带上鼠标键
7. 四种动作在动作列表中的描述文案
"""
import os
import sys
import time
import tkinter as tk
from tkinter import ttk
from threading import Lock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import auto_test_gui as ag


# --------------------------------------------------------------------------
# 纯逻辑测试（不需要 Tk）
# --------------------------------------------------------------------------
def test_build_insert_action_mapping():
    """四个选项必须映射到正确的动作字典"""
    click = ag.build_insert_action("left_click", 10, 20, 0.5)
    assert click == {"type": "click", "x": 10, "y": 20, "button": "left", "delay": 0.5}, click

    double = ag.build_insert_action("double_click", 10, 20, 0.5)
    assert double == {"type": "doubleclick", "x": 10, "y": 20, "button": "left", "delay": 0.5}, double

    right = ag.build_insert_action("right_click", 10, 20, 0.5)
    assert right == {"type": "click", "x": 10, "y": 20, "button": "right", "delay": 0.5}, right

    enter = ag.build_insert_action("enter", 10, 20, 0.5)
    assert enter == {"type": "enter", "x": 10, "y": 20, "delay": 0.5}, enter

    # 未知选项退化为默认的左键单击
    fallback = ag.build_insert_action("unknown", 1, 2, 0.0)
    assert fallback == {"type": "click", "x": 1, "y": 2, "button": "left", "delay": 0.0}, fallback

    # 默认选项就是单击
    default = ag.build_insert_action(ag.INSERT_ACTION_DEFAULT, 1, 2, 0.0)
    assert default["type"] == "click" and default["button"] == "left", default
    print("[OK] 四个选项分别构造出 click / doubleclick / 右键 click / enter 动作")


# --------------------------------------------------------------------------
# 对话框测试（使用真实 Tk）
# --------------------------------------------------------------------------
# 整个测试文件共用一个 Tk 根窗口：
# 反复创建/销毁 Tk 解释器会重新读取 init.tcl，在 Windows 上偶发
# "Can't find a usable init.tcl" 失败（与被测代码无关）。
_SHARED_ROOT = None


def get_shared_root():
    """获取（并缓存）共用的 Tk 根窗口"""
    global _SHARED_ROOT
    if _SHARED_ROOT is None:
        _SHARED_ROOT = tk.Tk()
        _SHARED_ROOT.withdraw()
    return _SHARED_ROOT


def destroy_shared_root():
    """释放共用的 Tk 根窗口（测试结束后调用）"""
    global _SHARED_ROOT
    if _SHARED_ROOT is not None:
        try:
            _SHARED_ROOT.destroy()
        except tk.TclError:
            pass
        _SHARED_ROOT = None


class DialogFake(ag.AutoTestGUI):
    """只构造"插入动作"对话框所需的字段"""

    def __init__(self):
        self.root = get_shared_root()
        self.actions_lock = Lock()
        self.actions = []
        self.in_dialog_operation = False
        self.action_tree = ttk.Treeview(self.root, columns=("action", "delay", "remark"))
        self.action_tree.pack()

    def _update_action_list(self, *a, **k):
        """跳过动作树刷新（避免依赖大量 UI 状态）"""

    def destroy(self):
        """共用根窗口由 destroy_shared_root() 统一释放，这里只清理自身状态"""
        self.actions = []
        self.in_dialog_operation = False


def walk_widgets(parent):
    """递归收集所有子控件"""
    result = [parent]
    for child in parent.winfo_children():
        result.extend(walk_widgets(child))
    return result


def find_dialog(root):
    """找到最近创建的 Toplevel 对话框"""
    for child in reversed(root.winfo_children()):
        if isinstance(child, tk.Toplevel):
            return child
    return None


def find_label(dialog, text):
    for w in walk_widgets(dialog):
        if isinstance(w, ttk.Label) and w.cget("text") == text:
            return w
    return None


def find_radios(dialog):
    return [w for w in walk_widgets(dialog) if isinstance(w, ttk.Radiobutton)]


def find_entries(dialog):
    """返回输入框（排除 Spinbox，它也是 ttk.Entry 的子类）"""
    return [w for w in walk_widgets(dialog)
            if isinstance(w, ttk.Entry) and not isinstance(w, ttk.Spinbox)]


def find_spinboxes(dialog):
    """返回数值输入框（Spinbox）"""
    return [w for w in walk_widgets(dialog) if isinstance(w, ttk.Spinbox)]


def radio_value(radio):
    """读取单选按钮当前绑定的变量值（ttk 不暴露 .variable 属性）"""
    var_name = radio.cget("variable")
    return radio.tk.globalgetvar(var_name)


def find_button(dialog, text):
    for w in walk_widgets(dialog):
        if isinstance(w, ttk.Button) and w.cget("text") == text:
            return w
    return None


def test_dialog_has_four_options_default_single_click():
    """对话框中必须有四个选项，且默认选中"单击\""""
    app = DialogFake()
    try:
        app.insert_action()
        app.root.update()
        dialog = find_dialog(app.root)
        assert dialog is not None, "未找到插入动作对话框"
        assert dialog.title() == "插入动作", f"对话框标题不对: {dialog.title()!r}"

        radios = find_radios(dialog)
        texts = [r.cget("text") for r in radios]
        assert texts == ["单击", "双击", "右键", "回车"], f"选项不对: {texts}"

        # 默认选中"单击"
        selected = [r.cget("text") for r in radios if radio_value(r) == r.cget("value")]
        assert selected == ["单击"], f"默认选项不是单击: {selected}"
        assert radio_value(radios[0]) == ag.INSERT_ACTION_DEFAULT, radio_value(radios[0])

        # 选项行必须位于"点击位置"行上方
        type_label = find_label(dialog, "动作类型:")
        pos_label = find_label(dialog, "点击位置:")
        assert type_label is not None, "缺少动作类型标签"
        assert pos_label is not None, "缺少点击位置标签"
        type_row = int(type_label.grid_info()["row"])
        pos_row = int(pos_label.grid_info()["row"])
        assert type_row < pos_row, f"选项行未位于点击位置行上方: {type_row} vs {pos_row}"

        dialog.destroy()
        app.in_dialog_operation = False
    finally:
        app.destroy()
    print("[OK] 对话框含 单击/双击/右键/回车 四个选项，默认选中单击，且位于点击位置行上方")


def test_dialog_inserts_action_matching_selection():
    """选择不同选项并确定后，插入的动作必须与选项一致"""
    app = DialogFake()
    try:
        expected = {
            "单击": {"type": "click", "button": "left"},
            "双击": {"type": "doubleclick", "button": "left"},
            "右键": {"type": "click", "button": "right"},
            "回车": {"type": "enter"},
        }
        for label, expect in expected.items():
            app.actions = []
            app.insert_action()
            app.root.update()
            dialog = find_dialog(app.root)

            radios = {r.cget("text"): r for r in find_radios(dialog)}
            assert label in radios, f"缺少选项 {label}"
            radios[label].invoke()          # 选中该选项
            app.root.update()

            entries = find_entries(dialog)
            assert len(entries) == 2, f"期望两个坐标输入框，实际 {len(entries)}"
            entries[0].insert(0, "123")
            entries[1].insert(0, "456")

            find_button(dialog, "确定").invoke()
            app.root.update()

            assert len(app.actions) == 1, f"{label}: 未插入动作 ({app.actions})"
            action = app.actions[0]
            assert action["type"] == expect["type"], f"{label}: 类型不对 {action}"
            assert action["x"] == 123 and action["y"] == 456, f"{label}: 坐标不对 {action}"
            if "button" in expect:
                assert action["button"] == expect["button"], f"{label}: 鼠标键不对 {action}"
            # 插入的动作必须记录相对延时，供下一次插入时作为默认值
            assert action["display_delay"] == action["delay"], f"{label}: display_delay 不对 {action}"

            dialog.destroy()
            app.in_dialog_operation = False
    finally:
        app.destroy()
    print("[OK] 四个选项分别插入对应的动作（类型/坐标/鼠标键均正确）")


def test_nested_click_edit_updates_delay():
    """循环组内编辑插入的动作时，延时必须真正写回动作（否则编辑结果丢失）"""
    app = DialogFake()
    try:
        nested = {"type": "enter", "x": 10, "y": 20, "delay": 1.5, "display_delay": 1.5}
        app.actions = [{
            "type": "loop_group",
            "loop_count": 3,
            "loop_actions": [nested],
        }]

        app._edit_nested_click_action(nested, "loop_0_0")
        app.root.update()
        dialog = find_dialog(app.root)
        assert dialog is not None, "未找到编辑对话框"

        spins = find_spinboxes(dialog)
        assert len(spins) == 1, f"期望一个延时输入框，实际 {len(spins)}"
        assert float(spins[0].get()) == 1.5, f"延时初值不对: {spins[0].get()}"

        spins[0].set(2.5)
        app.root.update()
        find_button(dialog, "确定").invoke()
        app.root.update()

        assert nested["delay"] == 2.5, f"延时未写回动作: {nested}"
        print("[OK] 循环组内编辑动作时延时正确写回")
    finally:
        app.destroy()


# --------------------------------------------------------------------------
# 回放测试
# --------------------------------------------------------------------------
class FakePyAutoGUI:
    """记录鼠标/键盘调用"""

    def __init__(self):
        self.FAILSAFE = True
        self._pos = (500, 500)
        self.clicks = []
        self.moves = []
        self.presses = []
        self.double_clicks = []

    def position(self):
        return self._pos

    def moveTo(self, x, y, duration=0):
        self.moves.append((x, y))

    def click(self, *a, **k):
        self.clicks.append(k.get("button", "left"))

    def doubleClick(self, *a, **k):
        self.double_clicks.append(k.get("button", "left"))

    def press(self, key, **k):
        self.presses.append(key)

    def scroll(self, *a, **k):
        pass

    def write(self, *a, **k):
        pass

    def hotkey(self, *a, **k):
        pass


class PlaybackFake(ag.AutoTestGUI):
    """跳过真实 __init__，只构造回放所需字段"""

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


class FakeRoot:
    def __init__(self):
        self.queue = []

    def after(self, delay_ms, func):
        self.queue.append(func)
        return len(self.queue)

    def flush(self):
        pending, self.queue = self.queue, []
        for func in pending:
            func()


class FakeVar:
    def __init__(self, value=None):
        self._value = value

    def get(self):
        return self._value

    def set(self, value):
        self._value = value


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


class FakeList:
    def __init__(self):
        self.items = []

    def delete(self, *a):
        self.items = []

    def insert(self, *a):
        self.items.append(a)

    def see(self, *a):
        pass


def run_playback(action, timeout=5.0):
    """执行一次回放，返回 FakePyAutoGUI 记录到的调用"""
    gui = FakePyAutoGUI()
    old_pyautogui = ag.pyautogui
    ag.pyautogui = gui
    try:
        app = PlaybackFake()
        app.actions = [action]
        app._start_playback_after_recording_stop()
        thread = app.playback_thread
        thread.join(timeout=timeout)
        assert not thread.is_alive(), "回放线程未退出"
        app.root.flush()
        return gui
    finally:
        ag.pyautogui = old_pyautogui


def test_enter_action_playback():
    """回车动作：先点击目标位置，再按下回车键"""
    gui = run_playback({"type": "enter", "x": 300, "y": 400, "delay": 0})
    assert gui.moves == [(300, 400)], f"未移动到目标位置: {gui.moves}"
    assert gui.clicks == ["left"], f"回车前应先单击目标位置: {gui.clicks}"
    assert gui.presses == ["enter"], f"未按下回车键: {gui.presses}"
    print("[OK] 回车动作回放时先点击目标位置再按下回车键")


def test_right_click_and_doubleclick_playback():
    """右键动作使用右键；双击动作带上鼠标键"""
    gui = run_playback({"type": "click", "x": 1, "y": 2, "button": "right", "delay": 0})
    assert gui.clicks == ["right"], f"右键动作未使用右键: {gui.clicks}"

    gui = run_playback({"type": "doubleclick", "x": 1, "y": 2, "button": "right", "delay": 0})
    assert gui.double_clicks == ["right"], f"双击动作未带上鼠标键: {gui.double_clicks}"
    print("[OK] 右键动作使用右键，双击动作带上鼠标键")


def test_describe_action_text():
    """四种动作在动作列表中的描述"""
    app = PlaybackFake()
    assert "单击 left" in app._describe_action({"type": "click", "x": 1, "y": 2, "button": "left"})
    assert "双击" in app._describe_action({"type": "doubleclick", "x": 1, "y": 2, "button": "left"})
    assert "单击 right" in app._describe_action({"type": "click", "x": 1, "y": 2, "button": "right"})
    assert app._describe_action({"type": "enter", "x": 1, "y": 2}) == "回车 at (1, 2)"
    print("[OK] 四种动作的列表描述文案正确")


def run():
    test_build_insert_action_mapping()
    test_describe_action_text()
    test_enter_action_playback()
    test_right_click_and_doubleclick_playback()
    test_dialog_has_four_options_default_single_click()
    test_dialog_inserts_action_matching_selection()
    test_nested_click_edit_updates_delay()
    destroy_shared_root()
    print("\n全部通过")


if __name__ == "__main__":
    run()
