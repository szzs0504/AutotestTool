"""集成测试：构造真实GUI，验证循环设置在保存/加载后的完整往返（含Treeview与控件状态）"""
import json
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import auto_test_gui as ag

# 绕过管理员检查（测试环境可能非管理员）
ag.is_admin = lambda: True

tmp = tempfile.mkdtemp(prefix="atg_int_")
newfile = os.path.join(tmp, "actions.json")
oldfile = os.path.join(tmp, "old.json")

with open(oldfile, "w", encoding="utf-8") as f:
    json.dump([{"type": "click", "x": 5, "y": 6, "button": "left", "delay": 0.3, "remark": "旧"}],
              f, ensure_ascii=False)

app = ag.AutoTestGUI()
app.root.withdraw()

ag.filedialog.asksaveasfilename = lambda **k: newfile
ag.filedialog.askopenfilename = lambda **k: newfile
ag.messagebox.showinfo = lambda *a, **k: None
ag.messagebox.showwarning = lambda *a, **k: None
errs = []
ag.messagebox.showerror = lambda *a, **k: errs.append(a)

failures = []


def check(cond, msg):
    print(("[OK] " if cond else "[FAIL] ") + msg)
    if not cond:
        failures.append(msg)


# --- 设置主界面循环参数 ---
app.loop_count.set(12)
app.loop_interval.set(4.5)
app.interval_random.set(False)
app.interval_min.set(1.0)
app.interval_max.set(6.0)

# --- 放入动作并保存 ---
app.actions = [
    {"type": "click", "x": 100, "y": 200, "button": "left", "delay": 0.4, "remark": "点击"},
    {"type": "input", "x": 110, "y": 210, "text": "hi", "delay": 0.2, "remark": "输入"},
]
app.save_actions()
check(not errs, f"保存无错误 ({errs})")

with open(newfile, encoding="utf-8") as f:
    saved = json.load(f)
check("settings" in saved and "actions" in saved, "文件包含 settings 与 actions 顶层键")
check(saved["settings"]["loop_count"] == 12, f"保存的次数正确: {saved['settings'].get('loop_count')}")
check(saved["settings"]["loop_interval"] == 4.5, "保存的固定间隔正确")
check(saved["settings"]["interval_min"] == 1.0 and saved["settings"]["interval_max"] == 6.0, "保存的随机范围正确")
check(saved["settings"]["interval_random"] is False, "保存的随机开关正确")
check(len(saved["actions"]) == 2, "保存了2个动作")

# --- 改变设置后加载，应被文件覆盖 ---
app.loop_count.set(1)
app.loop_interval.set(1.0)
app.interval_random.set(True)
app.interval_min.set(0.5)
app.interval_max.set(1.5)
app.actions = []

ag.filedialog.askopenfilename = lambda **k: newfile
app.load_actions()
check(not errs, f"加载新格式无错误 ({errs})")
check(app.loop_count.get() == 12, f"次数恢复为12: {app.loop_count.get()}")
check(app.loop_interval.get() == 4.5, f"固定间隔恢复为4.5: {app.loop_interval.get()}")
check(app.interval_random.get() is False, "随机开关恢复为关闭")
check(app.interval_min.get() == 1.0 and app.interval_max.get() == 6.0, "随机范围恢复")
check(len(app.actions) == 2, "动作已加载")
check(app.action_tree.get_children() != (), "Treeview 已刷新")

# --- 随机模式下的加载：间隔输入框应禁用 ---
app.interval_random.set(True)
app.loop_count.set(3)
app.loop_interval.set(2.0)
app.interval_min.set(0.2)
app.interval_max.set(3.3)
app.save_actions()
app.interval_random.set(False)
app.loop_count.set(1)
app.load_actions()
state = str(app.interval_entry.cget("state"))
check("disabled" in state, f"随机模式下间隔输入框禁用: {state}")
check(app.loop_count.get() == 3, f"次数恢复为3: {app.loop_count.get()}")
check(app.interval_min.get() == 0.2 and app.interval_max.get() == 3.3, "随机范围恢复")

# --- 加载旧格式文件：动作可读，设置保持当前值 ---
ag.filedialog.askopenfilename = lambda **k: oldfile
app.load_actions()
check(not errs, f"加载旧格式无错误 ({errs})")
check(len(app.actions) == 1, f"旧格式动作已加载: {len(app.actions)}")
check(app.actions[0]["remark"] == "旧", "旧格式动作内容正确")
check(app.loop_count.get() == 3, f"旧格式不覆盖设置，次数仍为3: {app.loop_count.get()}")

app.root.destroy()
print("\n" + ("全部通过" if not failures else f"失败 {len(failures)} 项: {failures}"))
sys.exit(1 if failures else 0)
