"""回归测试：动作文件保存/加载时循环设置的往返一致性

覆盖：
- save_actions() 写出 {version, settings, actions}
- load_actions() 从新格式恢复次数/间隔/随机开关/随机范围
- 旧格式（纯数组）仍可加载，且不覆盖当前设置
- 非法文件、非法 settings 字段被安全处理
"""
import json
import os
import sys
import tempfile
import tkinter as tk
from tkinter import ttk
from threading import Lock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import auto_test_gui as ag


class Fake(ag.AutoTestGUI):
    """跳过 __init__（需要管理员权限），只构造与保存/加载相关的字段"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.actions_lock = Lock()
        self.loop_count = tk.IntVar(value=1)
        self.loop_interval = tk.DoubleVar(value=1.0)
        self.interval_random = tk.BooleanVar(value=False)
        self.interval_min = tk.DoubleVar(value=0.5)
        self.interval_max = tk.DoubleVar(value=1.5)
        self.interval_entry = ttk.Entry(self.root, textvariable=self.loop_interval)
        self.status_var = tk.StringVar(value="")
        self.save_dir = tempfile.mkdtemp(prefix="atg_save_")
        self.actions = []
        self._update_action_list = lambda *a, **k: None   # 跳过树刷新
        self._update_loop_count_max = lambda: None

    def destroy(self):
        self.root.destroy()


def patch_dialogs(path):
    """把文件对话框/消息框替换为自动化桩"""
    ag.filedialog.asksaveasfilename = lambda **k: path
    ag.filedialog.askopenfilename = lambda **k: path
    ag.messagebox.showinfo = lambda *a, **k: None
    ag.messagebox.showwarning = lambda *a, **k: None
    ag.messagebox.showerror = lambda *a, **k: None


def collect_errors():
    errs = []
    ag.messagebox.showerror = lambda *a, **k: errs.append(a)
    return errs


def run():
    tmp = tempfile.mkdtemp(prefix="atg_test_")
    new_fmt = os.path.join(tmp, "new.json")
    old_fmt = os.path.join(tmp, "old.json")
    broken = os.path.join(tmp, "broken.json")

    with open(old_fmt, "w", encoding="utf-8") as f:
        json.dump([
            {"type": "click", "x": 1, "y": 2, "button": "left", "time": 0.0, "remark": "a"},
            {"type": "click", "x": 3, "y": 4, "button": "Button.right", "time": 2.0, "remark": "b"},
        ], f, ensure_ascii=False)

    with open(new_fmt, "w", encoding="utf-8") as f:
        json.dump({
            "version": 2,
            "settings": {
                "loop_count": 7, "loop_interval": 2.5, "interval_random": True,
                "interval_min": 0.3, "interval_max": 4.2,
            },
            "actions": [
                {"type": "click", "x": 1, "y": 2, "button": "left", "delay": 0.5, "time": 0.0},
            ],
        }, f, ensure_ascii=False)

    with open(broken, "w", encoding="utf-8") as f:
        f.write("not json at all")

    # ---- 1) 保存：真实 save_actions() 应写出带 settings 的新格式 ----
    app = Fake()
    app.actions = [{"type": "click", "x": 1, "y": 2, "button": "left", "delay": 0.2}]
    app.loop_count.set(5)
    app.loop_interval.set(3.5)
    app.interval_random.set(True)
    app.interval_min.set(1.25)
    app.interval_max.set(9.75)
    saved_to = os.path.join(tmp, "saved.json")
    patch_dialogs(saved_to)
    app.save_actions()
    with open(saved_to, encoding="utf-8") as f:
        saved = json.load(f)
    assert saved["settings"] == {
        "loop_count": 5, "loop_interval": 3.5, "interval_random": True,
        "interval_min": 1.25, "interval_max": 9.75,
    }, saved["settings"]
    assert saved["actions"] == app.actions
    assert saved["version"] == 2
    print("[OK] save_actions() 写出 {version, settings, actions}，设置值正确")

    # ---- 2) 加载新格式：真实 load_actions() 应恢复设置 ----
    app2 = Fake()
    patch_dialogs(new_fmt)
    errs = collect_errors()
    app2.load_actions()
    assert not errs, errs
    assert app2.loop_count.get() == 7, app2.loop_count.get()
    assert app2.loop_interval.get() == 2.5
    assert app2.interval_random.get() is True
    assert app2.interval_min.get() == 0.3
    assert app2.interval_max.get() == 4.2
    assert "disabled" in str(app2.interval_entry.cget("state"))
    assert len(app2.actions) == 1
    print("[OK] load_actions() 从新格式恢复循环设置，随机模式下间隔框已禁用")

    # ---- 3) 加载旧格式（纯数组）：动作可读、设置保持默认、button 兼容、delay 正确 ----
    app3 = Fake()
    patch_dialogs(old_fmt)
    errs = collect_errors()
    app3.load_actions()
    assert not errs, errs
    assert len(app3.actions) == 2, app3.actions
    assert app3.actions[1]["button"] == "right", app3.actions[1]["button"]
    assert app3.actions[0]["delay"] == 0.0, app3.actions[0]["delay"]
    assert app3.actions[1]["delay"] == 2.0, app3.actions[1]["delay"]
    assert app3.loop_count.get() == 1, "旧文件不应覆盖当前设置"
    assert app3.interval_random.get() is False
    print("[OK] 旧格式（纯数组）仍可加载；delay/button 兼容正确；设置不被覆盖")

    # ---- 4) 非法文件应弹错误而不是崩溃 ----
    app4 = Fake()
    patch_dialogs(broken)
    errs = collect_errors()
    app4.load_actions()
    assert len(errs) == 1, errs
    print("[OK] 非法文件弹出错误提示：", errs[0][1][:40])

    # ---- 5) 缺失/非法 settings 字段不应崩溃 ----
    app5 = Fake()
    patch_dialogs(saved_to)
    app5._apply_global_settings({"loop_count": 3})
    assert app5.loop_count.get() == 3
    app5._apply_global_settings({"loop_count": "abc"})
    app5._apply_global_settings(None)
    app5._apply_global_settings("not a dict")
    app5._apply_global_settings({"loop_interval": "x"})
    print("[OK] 缺失/非法 settings 字段被安全忽略")

    # ---- 6) 无动作时保存应给出警告 ----
    app6 = Fake()
    app6.actions = []
    warns = []
    patch_dialogs(saved_to)
    ag.messagebox.showwarning = lambda *a, **k: warns.append(a)
    app6.save_actions()
    assert len(warns) == 1, warns
    print("[OK] 空动作列表保存时给出警告")

    for a in (app, app2, app3, app4, app5, app6):
        a.destroy()
    print("\n全部通过")


if __name__ == "__main__":
    run()
