"""
Expert RC — Shear and torsion 抗剪验算自动化

读取 input_VN.xlsx「抗剪」sheet，按区域+编号分组：
  荷载表 Type / V / N（ULS/ALS）
  选项 shear、Seismic、Critical、Column；不选 torsion
  箍筋 n1/d_bar1（有 n2/d_bar2 时按面积折算并入 n1，仅填第一组）
  尺寸 b、h、d（仅 d1_cm → d）
  荷载表仅约 4 行可见：第 1～4 行直接填，第 5 行起翻页（勿与 NM 9 行逻辑混用）
  Calculate → Note → notes/{区域}_shear.rtf（同区域多编号时追加编号）
  结果汇总 → result_shear_summary.xlsx（sheet「结果汇总」）

用法：
  cd expert/expert_shear
  python expert_try_shear.py
"""
import win32gui
import win32con
import win32api
import win32process
import ctypes
from ctypes import wintypes
import json
import sys
import time
import pandas as pd
import shutil
import os
import glob
import re
import math
import threading
from collections import defaultdict
from typing import Dict, List, Optional

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ===================== 配置 =====================
TARGET_WINDOW = "EXPERT RC - Shear and torsion"
INPUT_EXCEL = os.path.join(SCRIPT_DIR, "input_VN.xlsx")
INPUT_SHEET = "抗剪"
OUTPUT_EXCEL = os.path.join(SCRIPT_DIR, "result_shear_summary.xlsx")

# 结果汇总表列（单独 Excel「结果汇总」sheet）
RESULT_SUMMARY_COLUMNS = [
    "区域", "编号",
    "b_cm", "h_cm", "d_cm",
    "n1", "d_bar1", "n2", "d_bar2",
    "荷载组数", "荷载类型", "V_control_kN", "N_control_kN",
    "St_cm", "St_max_cm", "计算状态", "结果摘要", "RTF路径",
]

# ---------- 等待时间（调大=更慢更稳，调小=更快但易出错）----------
SLEEP_SHORT = 0.05         # 控件点击、勾选 Seismic 等
SLEEP_CALC = 0.1          # Calculate 后额外稳定等待（弹框关闭后再等）
SLEEP_CALC_MAX = 5.0      # Calculate 后最长等待（轮询 rcba 直至关闭或超时）
SLEEP_NOTE = 1.5           # Note 后等待 RTF 生成
SLEEP_CELL_CLICK = 0.1    # 荷载表物理/模拟点击单元格后
SLEEP_CELL_SCROLL = 0.25    # Down 翻页后（SendInput 真实按键，需稍长）
SLEEP_COMMIT_CELL = 0.1    # 单格 Enter 提交后
SLEEP_TYPE_CELL = 0.01      # F2 进入单元格编辑后
SLEEP_GRID_CLEAR = 0.05    # 清空荷载（点行号 / Delete 每步）
SLEEP_LOAD_START = 0.05     # 清空荷载完毕、开始填入前
SLEEP_RCBA_POLL = 0.15     # 轮询 rcba 对话框间隔
SLEEP_KEYBOARD_SESSION = 0.2  # 荷载填表前 AttachThreadInput 抢焦点后
SLEEP_KEY_CHAR = 0.025     # 逐字符 WM_CHAR 间隔（grid_send_text）
SLEEP_KEY_VK = 0.04        # 单键 SendInput 间隔（send_vk）
SLEEP_STIRRUP_N2 = 0.45   # 填 n2 后等待第二组 φ 下拉框启用
SLEEP_STIRRUP_FIELD = 0.12  # 箍筋坐标点击后等待
SLEEP_SAFE_CLICK = 0.15    # safe_click 物理鼠标点击后
SLEEP_FIND_RTF = 0.3       # find_rc_note_file 轮询间隔
# Expert 默认输出路径；Note 按钮生成的 rc_note.rtf 在此
EXPERT_RTF_PATH = os.path.join(
    os.path.expanduser("~"), "Documents", "Autodesk", "Output", "rc_note.rtf"
)
# 按计算编号另存 RTF 的目标文件夹（可改成任意路径）
OUTPUT_RTF_DIR = os.path.join(SCRIPT_DIR, "notes")

# 荷载表坐标（相对 GXWND 左上角）；运行 calibrate_load.py 自动写入 load_coords.json
LOAD_COORDS_FILE = os.path.join(SCRIPT_DIR, "load_coords_shear.json")
LOAD_ROW_INDEX_COL_X = 20
LOAD_TYPE_COL_X = 69
LOAD_V_COL_X = 120
LOAD_N_COL_X = 171
LOAD_EXPAND_BTN_X = 250
LOAD_EXPAND_BTN_Y = 120
LOAD_FIRST_ROW_Y = 2
LOAD_ROW_HEIGHT = 19.8
LOAD_CLEAR_TIMES = 12
# 抗剪 Loads 表界面一次约显示 4 行（仅影响翻页坐标，不限制荷载组数）
SHEAR_LOAD_VISIBLE_ROWS = 4
# 第 5 行起翻页：点第 4 行 Load type + Down；填表 Y 取第 3 行（与双轴/NM 相同，最底行易点到滚动条）
SHEAR_LOAD_SCROLL_CLICK_ROW = SHEAR_LOAD_VISIBLE_ROWS - 1
SHEAR_LOAD_OVERFLOW_FILL_ROW = max(0, SHEAR_LOAD_VISIBLE_ROWS - 1)


def apply_load_coords():
    """优先读取 load_coords_shear.json（含荷载表与箍筋坐标）。"""
    global LOAD_ROW_INDEX_COL_X, LOAD_TYPE_COL_X, LOAD_V_COL_X, LOAD_N_COL_X
    global LOAD_EXPAND_BTN_X, LOAD_EXPAND_BTN_Y
    global LOAD_FIRST_ROW_Y, LOAD_ROW_HEIGHT
    global STIRRUP_COORDS
    if not os.path.isfile(LOAD_COORDS_FILE):
        STIRRUP_COORDS = {}
        return False
    with open(LOAD_COORDS_FILE, encoding="utf-8") as f:
        d = json.load(f)
    LOAD_ROW_INDEX_COL_X = int(d["LOAD_ROW_INDEX_COL_X"])
    LOAD_TYPE_COL_X = int(d["LOAD_TYPE_COL_X"])
    LOAD_FIRST_ROW_Y = int(d["LOAD_FIRST_ROW_Y"])
    LOAD_ROW_HEIGHT = int(d["LOAD_ROW_HEIGHT"])
    LOAD_V_COL_X = int(d.get("LOAD_V_COL_X", LOAD_TYPE_COL_X + 51))
    LOAD_N_COL_X = int(d.get("LOAD_N_COL_X", LOAD_V_COL_X + 51))
    LOAD_EXPAND_BTN_X = int(d.get("LOAD_EXPAND_BTN_X", LOAD_N_COL_X + 70))
    LOAD_EXPAND_BTN_Y = int(d.get("LOAD_EXPAND_BTN_Y", LOAD_FIRST_ROW_Y + 80))
    STIRRUP_COORDS = {
        k: int(d[k])
        for k in d
        if k.startswith("STIRRUP_")
    }
    return True


STIRRUP_COORDS: Dict[str, int] = {}
apply_load_coords()

EDIT_CLASSES = {"Edit", "TEdit", "TMaskEdit"}
RICHEDIT_CLASSES = {"RichEdit", "RichEdit20A", "RichEdit20W", "RichEdit50W"}
BUTTON_CLASS = {"Button", "TButton"}
SPIN_CLASSES = {"TSpinEdit", "SpinEdit"}
UPDOWN_CLASS = "msctls_updown32"
UDM_GETBUDDY = 0x046E
WM_NEXTDLGCTL = 0x0280
COMBOBOX_CLASS = "ComboBox"
DIM_CTRL_IDS = {"edit_b": 1227, "edit_h": 1228, "edit_d1": 1232, "edit_d2": 1233}


class _GUITHREADINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", ctypes.c_ulong),
        ("flags", ctypes.c_ulong),
        ("hwndActive", wintypes.HWND),
        ("hwndFocus", wintypes.HWND),
        ("hwndCapture", wintypes.HWND),
        ("hwndMenuOwner", wintypes.HWND),
        ("hwndMoveSize", wintypes.HWND),
        ("hwndCaret", wintypes.HWND),
        ("rcCaret", wintypes.RECT),
    ]


def get_expert_focus(main_hw) -> Optional[int]:
    """跨线程读取 Expert 线程焦点（GetFocus 在自动化脚本中通常无效）。"""
    if not is_valid_hwnd(main_hw):
        return None
    tid, _ = win32process.GetWindowThreadProcessId(main_hw)
    gti = _GUITHREADINFO()
    gti.cbSize = ctypes.sizeof(_GUITHREADINFO)
    if not ctypes.windll.user32.GetGUIThreadInfo(tid, ctypes.byref(gti)):
        return None
    focus = gti.hwndFocus
    return focus if is_valid_hwnd(focus) else None


# ==================================================


def get_expert_hwnd():
    hwnds = []
    win32gui.EnumWindows(
        lambda hw, arr: arr.append(hw) or True
        if win32gui.IsWindowVisible(hw) and TARGET_WINDOW in win32gui.GetWindowText(hw)
        else True,
        hwnds,
    )
    if not hwnds:
        raise RuntimeError(f"未找到 Expert 窗口：{TARGET_WINDOW}")
    return hwnds[0]


def is_valid_hwnd(hwnd):
    try:
        return bool(hwnd) and win32gui.IsWindow(hwnd)
    except Exception:
        return False


def refresh_expert_hwnd(main_hw=None):
    """句柄失效时按窗口标题重新查找 Expert 主窗口"""
    if main_hw and is_valid_hwnd(main_hw):
        try:
            if TARGET_WINDOW in win32gui.GetWindowText(main_hw):
                return main_hw
        except Exception:
            pass
    return get_expert_hwnd()


def ensure_expert_visible(main_hw=None):
    """恢复最小化窗口；返回当前有效的主窗口句柄"""
    main_hw = refresh_expert_hwnd(main_hw)
    if win32gui.IsIconic(main_hw):
        win32gui.ShowWindow(main_hw, win32con.SW_RESTORE)
        time.sleep(0.4)
        main_hw = refresh_expert_hwnd()
    if not win32gui.IsWindowVisible(main_hw):
        raise RuntimeError("Expert 窗口不可见，请先打开并置于 Design 页")
    return main_hw


def get_design_page_hwnd(main_hw):
    found = []

    def walk(hwnd):
        try:
            if win32gui.GetWindowText(hwnd) == "Design":
                found.append(hwnd)
            win32gui.EnumChildWindows(hwnd, lambda c, _: walk(c) or True, None)
        except Exception:
            pass

    walk(main_hw)
    if not found:
        raise RuntimeError("未找到 Design 页，请手动切换到 Design")
    return found[0]


def get_design_page_rect(main_hw):
    return win32gui.GetWindowRect(get_design_page_hwnd(main_hw))


def has_stirrup_coords() -> bool:
    needed = [f"STIRRUP_{k}_X" for k in ("N1", "D1", "N2", "D2")]
    needed += [f"STIRRUP_{k}_Y" for k in ("N1", "D1", "N2", "D2")]
    return all(k in STIRRUP_COORDS for k in needed)


def stirrup_screen_point(main_hw, key: str) -> tuple:
    dl, dt, _, _ = get_design_page_rect(main_hw)
    return (
        dl + STIRRUP_COORDS[f"STIRRUP_{key}_X"],
        dt + STIRRUP_COORDS[f"STIRRUP_{key}_Y"],
    )


def refresh_stirrup_hwnds(main_hw, ctrls) -> bool:
    """填箍筋前重新扫描 n1/φ1（仅第一组，不定位 n2/φ2）。"""
    rows, _ = get_page_rows(main_hw)
    n1, d1, _, _, plus = find_stirrup_controls(rows, ctrls["grid_rect"])
    ctrls["n1_hw"], ctrls["d1_cb"] = n1, d1
    ctrls["n2_hw"], ctrls["d2_cb"] = None, None
    ctrls["stirrup_plus"] = plus
    return bool(n1 and d1)


def _hwnd_click_point(hwnd, combo: bool = False) -> tuple:
    """点击控件左侧（避开 Spinner/Combo 右侧下拉箭头）。"""
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width = max(right - left, 1)
    if combo:
        x = left + min(width // 3, max(width - 14, 8))
    else:
        x = left + min(width // 3, max(width - 10, 6))
    y = (top + bottom) // 2
    return x, y


def _click_hwnd(main_hw, hwnd, combo: bool = False):
    if not is_valid_hwnd(hwnd):
        return
    x, y = _hwnd_click_point(hwnd, combo=combo)
    safe_click(main_hw, x, y)
    time.sleep(SLEEP_STIRRUP_FIELD)
    try:
        win32gui.SetForegroundWindow(main_hw)
        win32gui.SetFocus(hwnd)
    except Exception:
        pass
    time.sleep(SLEEP_SHORT)


def _resolve_click_target(hwnd):
    """点击落到 Button（Spinner 箭头）时，向上找同级 Edit。"""
    if not is_valid_hwnd(hwnd):
        return hwnd
    cls = win32gui.GetClassName(hwnd)
    if cls in EDIT_CLASSES or cls == COMBOBOX_CLASS:
        return hwnd
    if cls in BUTTON_CLASS:
        parent = win32gui.GetParent(hwnd)
        if is_valid_hwnd(parent):
            edits = []
            def cb(hw, _):
                if win32gui.GetClassName(hw) in EDIT_CLASSES:
                    edits.append(hw)
                return True
            win32gui.EnumChildWindows(parent, cb, None)
            if edits:
                edits.sort(key=lambda h: win32gui.GetWindowRect(h)[0])
                return edits[0]
    return hwnd


def _hwnd_after_screen_click(main_hw, x, y):
    """物理点击屏幕坐标，返回该位置的 Edit/ComboBox 句柄。"""
    safe_click(main_hw, int(x), int(y))
    time.sleep(SLEEP_STIRRUP_FIELD)
    try:
        win32gui.SetForegroundWindow(main_hw)
    except Exception:
        pass
    pt = (int(x), int(y))
    hw = win32gui.WindowFromPoint(pt)
    hw = _resolve_click_target(hw)
    if win32gui.GetClassName(hw) in EDIT_CLASSES:
        parent = win32gui.GetParent(hw)
        if parent and win32gui.GetClassName(parent) == COMBOBOX_CLASS:
            return parent
        return hw
    if win32gui.GetClassName(hw) == COMBOBOX_CLASS:
        return hw
    focus = get_expert_focus(main_hw)
    if is_valid_hwnd(focus):
        fcls = win32gui.GetClassName(focus)
        if fcls in EDIT_CLASSES or fcls == COMBOBOX_CLASS:
            return focus
    return hw


def _combo_list_items(hwnd, open_drop: bool = False) -> List[str]:
    if not is_valid_hwnd(hwnd):
        return []
    if open_drop:
        try:
            win32gui.SendMessage(hwnd, win32con.CB_SHOWDROPDOWN, True, 0)
            time.sleep(0.1)
        except Exception:
            pass
    count = win32gui.SendMessage(hwnd, win32con.CB_GETCOUNT, 0, 0)
    items = []
    for i in range(max(int(count), 0)):
        buf = ctypes.create_unicode_buffer(64)
        win32gui.SendMessage(hwnd, win32con.CB_GETLBTEXT, i, buf)
        items.append(buf.value.strip())
    if open_drop:
        try:
            win32gui.SendMessage(hwnd, win32con.CB_SHOWDROPDOWN, False, 0)
        except Exception:
            pass
    return items


def _hwnd_is_descendant_of(hwnd, ancestor) -> bool:
    if not hwnd or not ancestor:
        return False
    p = hwnd
    while p:
        if p == ancestor:
            return True
        try:
            p = win32gui.GetParent(p)
        except Exception:
            break
    return False


def _combo_inner_edit_hwnds(combo_hwnd) -> set:
    inner = set()
    if not is_valid_hwnd(combo_hwnd):
        return inner
    def cb(hw, _):
        if win32gui.GetClassName(hw) in EDIT_CLASSES:
            inner.add(hw)
        return True
    try:
        win32gui.EnumChildWindows(combo_hwnd, cb, None)
    except Exception:
        pass
    return inner


def _row_center_x(row) -> int:
    return row["left"] + row["width"] // 2


def _row_right(row) -> int:
    return row["left"] + row["width"]


def _edit_inside_combo_box(edit_row, combo_row) -> bool:
    cx = _row_center_x(edit_row)
    cy = edit_row["top"] + edit_row["height"] // 2
    return (
        combo_row["left"] - 2 <= cx <= _row_right(combo_row) + 2
        and abs(edit_row["top"] - combo_row["top"]) <= 16
    )


def _filter_combo_inner_edits(edits, combos) -> list:
    """去掉 ComboBox 内部 Edit（会被误当成 n2）。"""
    inner = set()
    for c in combos:
        inner |= _combo_inner_edit_hwnds(c["hwnd"])
    out = []
    for e in edits:
        if e["hwnd"] in inner:
            continue
        if any(_hwnd_is_descendant_of(e["hwnd"], c["hwnd"]) for c in combos):
            continue
        if any(_edit_inside_combo_box(e, c) for c in combos):
            continue
        out.append(e)
    return out


def _is_d1_combo_edit(hwnd, d1_cb) -> bool:
    """φ1 ComboBox 内部 Edit（句柄或屏幕坐标落在 φ1 框内）。"""
    if not is_valid_hwnd(hwnd) or not d1_cb:
        return False
    if hwnd in _combo_inner_edit_hwnds(d1_cb):
        return True
    if _hwnd_is_descendant_of(hwnd, d1_cb):
        return True
    d1l, d1t, d1r, d1b = win32gui.GetWindowRect(d1_cb)
    l, t, r, b = win32gui.GetWindowRect(hwnd)
    cx, cy = (l + r) // 2, (t + b) // 2
    return (
        d1l - 4 <= cx <= d1r + 4
        and d1t - 8 <= cy <= d1b + 8
    )


def _n2_min_center_x(pl: int, d1_row) -> int:
    """n2 中心 x 下限：必须在 φ1 右缘之外（真 n2≈613，φ1 内 Edit≈532）。"""
    if d1_row:
        return max(pl + 4, _row_right(d1_row) + 2)
    return pl + 4


def _resolve_n2_edit(edits, pl: int, d1_row, combos) -> Optional[int]:
    """n2 必须在「+」右侧且 entirely 在 φ1 右缘之外。"""
    inner = set()
    for c in combos:
        inner |= _combo_inner_edit_hwnds(c["hwnd"])
    min_x = _n2_min_center_x(pl, d1_row)
    cands = []
    for e in edits:
        if e["hwnd"] in inner:
            continue
        if d1_row and _is_d1_combo_edit(e["hwnd"], d1_row["hwnd"]):
            continue
        if any(_edit_inside_combo_box(e, c) for c in combos):
            continue
        if _row_center_x(e) <= min_x:
            continue
        cands.append(e)
    if not cands:
        return None
    cands.sort(key=lambda x: x["left"])
    return cands[0]["hwnd"]


def _resolve_d2_combo(combos, pl: int, d1_row, pt: int, rows) -> Optional[int]:
    """φ2 为「+」右侧、最靠右的箍筋直径 Combo（禁用时也识别）。"""
    d1_right = _row_right(d1_row) if d1_row else pl
    min_x = max(pl + 8, d1_right + 5)
    cands = [c for c in combos if _row_center_x(c) > min_x]
    if not cands:
        for r in rows:
            if r["class"] != COMBOBOX_CLASS:
                continue
            if abs(r["top"] - pt) > 22:
                continue
            if _row_center_x(r) <= min_x:
                continue
            if d1_row and r["hwnd"] == d1_row["hwnd"]:
                continue
            cands.append(r)
    if not cands:
        return None
    bar = [c for c in cands if _is_rebar_diameter_combo(c["hwnd"])]
    pick = (bar or cands)
    pick.sort(key=lambda x: x["left"])
    return pick[-1]["hwnd"]


def _pick_rebar_combo(combo_rows):
    bar = [c for c in combo_rows if _is_rebar_diameter_combo(c["hwnd"])]
    return (bar or combo_rows)[0]["hwnd"] if combo_rows else None


def _is_rebar_diameter_combo(hwnd) -> bool:
    """区分箍筋直径下拉（6/8/10…）与倾角下拉（45/90）。"""
    items = _combo_list_items(hwnd, open_drop=True)
    if len(items) < 4:
        return False
    nums = []
    for it in items:
        m = re.search(r"(\d+)", it)
        if m:
            nums.append(int(m.group(1)))
    if not nums:
        return False
    if set(nums).issubset({45, 90}):
        return False
    return any(n <= 32 for n in nums)


def _combo_edit_child(hwnd):
    found = []
    def cb(hw, _):
        if win32gui.GetClassName(hw) in EDIT_CLASSES:
            found.append(hw)
        return True
    try:
        win32gui.EnumChildWindows(hwnd, cb, None)
    except Exception:
        pass
    return found[0] if found else None


def _read_combo_text(hwnd) -> str:
    if not is_valid_hwnd(hwnd):
        return ""
    child = _combo_edit_child(hwnd)
    if child:
        t = _read_edit_text(child)
        if t:
            return t
    sel = win32gui.SendMessage(hwnd, win32con.CB_GETCURSEL, 0, 0)
    if sel >= 0:
        buf = ctypes.create_unicode_buffer(32)
        win32gui.SendMessage(hwnd, win32con.CB_GETLBTEXT, sel, buf)
        if buf.value.strip():
            return buf.value.strip()
    return win32gui.GetWindowText(hwnd).strip()


def _combo_value_matches(got: str, expected: str) -> bool:
    if not got:
        return False
    bare = re.sub(r"^[φΦøØ]\s*", "", got).strip()
    return bare == expected or expected in got or got == expected


def _notify_combo_selchange(hwnd):
    parent = win32gui.GetParent(hwnd)
    if not parent:
        return
    ctrl_id = win32gui.GetDlgCtrlID(hwnd)
    win32gui.SendMessage(
        parent,
        win32con.WM_COMMAND,
        win32api.MAKELONG(ctrl_id, win32con.CBN_SELCHANGE),
        hwnd,
    )


def _stirrup_set_combo(main_hw, hwnd, value, label: str) -> bool:
    if not is_valid_hwnd(hwnd):
        print(f"  ✗ {label} 句柄无效")
        return False
    text = str(int(value))
    items = _combo_list_items(hwnd, open_drop=True)
    if items:
        print(f"  → {label} 下拉 {len(items)} 项: {', '.join(items[:10])}{'…' if len(items) > 10 else ''}")

    _click_hwnd(main_hw, hwnd, combo=True)

    # 1) CB_SETCURSEL
    idx = -1
    for i, item in enumerate(items):
        bare = re.sub(r"^[φΦøØ]\s*", "", item).strip()
        if item == text or bare == text:
            idx = i
            break
    if idx < 0:
        idx = win32gui.SendMessage(hwnd, win32con.CB_FINDSTRINGEXACT, -1, text)
    if idx < 0:
        idx = win32gui.SendMessage(hwnd, win32con.CB_FINDSTRING, -1, text)
    if idx >= 0:
        win32gui.SendMessage(hwnd, win32con.CB_SETCURSEL, idx, 0)
        _notify_combo_selchange(hwnd)
        time.sleep(0.1)
        if _combo_value_matches(_read_combo_text(hwnd), text):
            print(f"  → {label}={text} ✓")
            return True

    # 2) 子 Edit WM_SETTEXT
    child = _combo_edit_child(hwnd)
    if child:
        set_edit_text(child, text)
        _notify_combo_selchange(hwnd)
        time.sleep(0.1)
        if _combo_value_matches(_read_combo_text(hwnd), text):
            print(f"  → {label}={text} ✓ (子Edit)")
            return True

    # 3) 键盘：点下拉箭头 → 输入 → Enter
    with expert_keyboard_session(main_hw):
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        safe_click(main_hw, right - 8, (top + bottom) // 2)
        time.sleep(0.15)
        send_ctrl_a()
        send_text(text)
        send_vk(win32con.VK_RETURN)
        time.sleep(0.12)
    _notify_combo_selchange(hwnd)
    got = _read_combo_text(hwnd)
    if _combo_value_matches(got, text):
        print(f"  → {label}={text} ✓ (键盘)")
        return True

    print(f"  ✗ {label} 期望 {text}，读回 {got!r}")
    return False


def _notify_edit_change(hwnd):
    parent = win32gui.GetParent(hwnd)
    if not parent:
        return
    ctrl_id = win32gui.GetDlgCtrlID(hwnd)
    win32gui.SendMessage(
        parent, win32con.WM_COMMAND,
        win32api.MAKELONG(ctrl_id, win32con.EN_CHANGE), hwnd,
    )
    win32gui.SendMessage(
        parent, win32con.WM_COMMAND,
        win32api.MAKELONG(ctrl_id, win32con.EN_KILLFOCUS), hwnd,
    )


def _stirrup_set_edit(main_hw, hwnd, value, label: str, d1_cb=None, results_log_hw=None) -> bool:
    if not is_valid_hwnd(hwnd):
        print(f"  ✗ {label} 句柄无效")
        return False
    if results_log_hw and hwnd == results_log_hw:
        print(f"  ✗ {label} 句柄是 Results 输出框，已跳过")
        return False
    if d1_cb and (_hwnd_is_descendant_of(hwnd, d1_cb) or hwnd in _combo_inner_edit_hwnds(d1_cb)):
        print(f"  ✗ {label} 句柄是 φ1 内部 Edit，已跳过")
        return False
    _click_hwnd(main_hw, hwnd, combo=False)
    try:
        win32gui.SendMessage(hwnd, win32con.EM_SETSEL, 0, -1)
    except Exception:
        pass
    set_edit_text(hwnd, int(value))
    _notify_edit_change(hwnd)
    time.sleep(0.1)
    got = _read_edit_text(hwnd)
    try:
        ok = int(float(str(got).replace(",", "."))) == int(value)
    except (TypeError, ValueError):
        ok = str(got).strip() == str(int(value))
    if ok:
        print(f"  → {label}={int(value)} ✓")
    else:
        print(f"  ✗ {label} 期望 {int(value)}，读回 {got!r}")
    if ok and d1_cb and label == "n2":
        d1_got = _read_combo_text(d1_cb)
        bare = re.sub(r"^[φΦøØ]\s*", "", d1_got).strip()
        if bare == str(int(value)):
            print(f"  ✗ n2={int(value)} 误写入 φ1（φ1 现为 {d1_got!r}）")
            return False
    return ok


def _log_stirrup_hwnds(ctrls, dynamic: bool = False):
    parts = []
    for key, label in (
        ("n1_hw", "n1"), ("d1_cb", "φ1"),
        ("n2_hw", "n2"), ("d2_cb", "φ2"),
    ):
        hw = ctrls.get(key)
        if is_valid_hwnd(hw):
            cls = win32gui.GetClassName(hw)
            cid = win32gui.GetDlgCtrlID(hw)
            l, t, r, b = win32gui.GetWindowRect(hw)
            tag = label
            if dynamic and key in ("n2_hw", "d2_cb"):
                tag = f"{label}✓"
            parts.append(f"{tag}={cls}#{cid}@x{(l + r) // 2}")
        elif key in ("n2_hw", "d2_cb") and not dynamic:
            parts.append(f"{label}=待定位")
        else:
            parts.append(f"{label}=?")
    print(f"  → 箍筋控件: {', '.join(parts)}")


def _is_valid_n2_hwnd(
    n2_hw, n1_hw, d1_cb, plus_anchor=None, results_log_hw=None,
) -> bool:
    if not is_valid_hwnd(n2_hw):
        return False
    if results_log_hw and n2_hw == results_log_hw:
        return False
    if n1_hw and n2_hw == n1_hw:
        return False
    if d1_cb and (
        _is_d1_combo_edit(n2_hw, d1_cb)
    ):
        return False
    cls = win32gui.GetClassName(n2_hw)
    if cls in RICHEDIT_CLASSES:
        return False
    if cls not in EDIT_CLASSES and cls not in SPIN_CLASSES:
        return False
    l, t, r, b = win32gui.GetWindowRect(n2_hw)
    w, h = r - l, b - t
    if w > 120 or h > 50 or w < 10 or h < 10:
        return False
    if plus_anchor:
        pl = plus_anchor["left"]
        if (l + r) // 2 <= pl + 4:
            return False
    if d1_cb:
        d1l, _, d1r, _ = win32gui.GetWindowRect(d1_cb)
        if (l + r) // 2 <= d1r + 2:
            return False
    if n1_hw:
        n1l, n1t, n1r, n1b = win32gui.GetWindowRect(n1_hw)
        if abs(t - n1t) > 14:
            return False
        if abs(h - (n1b - n1t)) > 12:
            return False
        if (l + r) // 2 <= (n1l + n1r) // 2 + 6:
            return False
    return True


def _log_stirrup_row_candidates(rows, n1_hw, plus_anchor, label: str, d1_cb=None, results_log_hw=None):
    if not plus_anchor:
        return
    n1t = win32gui.GetWindowRect(n1_hw)[1]
    parts = []
    for r in rows:
        if r["class"] not in EDIT_CLASSES and r["class"] not in SPIN_CLASSES:
            continue
        if not _is_valid_n2_hwnd(
            r["hwnd"], n1_hw, d1_cb, plus_anchor, results_log_hw,
        ):
            continue
        cx = r["left"] + r["width"] // 2
        parts.append(f"{r['class']}#{r['ctrl_id']}@x{cx}")
    if parts:
        print(f"  → {label} 候选: {', '.join(parts[:8])}")


def _log_hwnd(label, hwnd):
    if not is_valid_hwnd(hwnd):
        print(f"  → {label}=?")
        return
    l, _, r, _ = win32gui.GetWindowRect(hwnd)
    print(
        f"  → {label}={win32gui.GetClassName(hwnd)}"
        f"#{win32gui.GetDlgCtrlID(hwnd)}@x{(l + r) // 2}"
    )


def _tab_once():
    send_vk(win32con.VK_TAB)
    time.sleep(0.08)


def collect_controls_deep(parent):
    """递归枚举子控件（Spinner 内 Edit 等嵌套控件需要）。"""
    rows = []

    def walk(hwnd):
        try:
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            rows.append({
                "hwnd": hwnd,
                "class": win32gui.GetClassName(hwnd),
                "text": win32gui.GetWindowText(hwnd),
                "ctrl_id": win32gui.GetDlgCtrlID(hwnd),
                "top": top, "left": left,
                "width": right - left, "height": bottom - top,
            })
            win32gui.EnumChildWindows(hwnd, lambda c, _: walk(c) or True, None)
        except Exception:
            pass

    walk(parent)
    rows.sort(key=lambda r: (r["top"], r["left"]))
    return rows


def _get_design_rows_deep(main_hw):
    try:
        design_hw = get_design_page_hwnd(main_hw)
        return collect_controls_deep(design_hw)
    except RuntimeError:
        return collect_all_controls(main_hw)


def _edit_from_spin_hwnd(hwnd):
    cls = win32gui.GetClassName(hwnd)
    if cls in EDIT_CLASSES:
        return hwnd
    if cls in SPIN_CLASSES:
        child = _combo_edit_child(hwnd)
        return child if is_valid_hwnd(child) else None
    return None


def _find_n2_edit_deep(rows, n1_hw, d1_cb, plus_anchor, results_log_hw=None):
    """深度扫描「+」右侧、与 n1 同行的 Spinner Edit（排除 Results 大文本框）。"""
    if not is_valid_hwnd(n1_hw):
        return None
    n1l, n1t, n1r, n1b = win32gui.GetWindowRect(n1_hw)
    pl = plus_anchor["left"] if plus_anchor else n1l
    if d1_cb:
        min_x = max(pl + 4, win32gui.GetWindowRect(d1_cb)[2] + 2)
    else:
        min_x = max(pl + 4, n1r + 2)
    inner = set()
    for r in rows:
        if r["class"] == COMBOBOX_CLASS:
            inner |= _combo_inner_edit_hwnds(r["hwnd"])

    # Spinner：「+」右侧第一个 UpDown 的 buddy 即为 n2
    for r in rows:
        if r["class"] != UPDOWN_CLASS:
            continue
        if abs(r["top"] - n1t) > 8:
            continue
        if r["left"] + r["width"] // 2 <= min_x:
            continue
        buddy = win32gui.SendMessage(r["hwnd"], UDM_GETBUDDY, 0, 0)
        if is_valid_hwnd(buddy) and buddy != n1_hw:
            if _is_valid_n2_hwnd(
                buddy, n1_hw, d1_cb, plus_anchor, results_log_hw,
            ):
                return buddy

    cands = []

    def try_add(hwnd, left, top):
        if not is_valid_hwnd(hwnd):
            return
        if hwnd == n1_hw or hwnd in inner:
            return
        if d1_cb and _hwnd_is_descendant_of(hwnd, d1_cb):
            return
        if not _is_valid_n2_hwnd(
            hwnd, n1_hw, d1_cb, plus_anchor, results_log_hw,
        ):
            return
        cx = left + (win32gui.GetWindowRect(hwnd)[2] - win32gui.GetWindowRect(hwnd)[0]) // 2
        if cx < min_x:
            return
        cands.append((cx, hwnd))

    for r in rows:
        if abs(r["top"] - n1t) > 8:
            continue
        cx = r["left"] + r["width"] // 2
        if cx < min_x:
            continue
        if results_log_hw and r["hwnd"] == results_log_hw:
            continue
        cls = r["class"]
        if cls in EDIT_CLASSES:
            try_add(r["hwnd"], r["left"], r["top"])
        elif cls in SPIN_CLASSES:
            ed = _edit_from_spin_hwnd(r["hwnd"])
            if ed:
                try_add(ed, r["left"], r["top"])
        elif cls == UPDOWN_CLASS:
            buddy = win32gui.SendMessage(r["hwnd"], UDM_GETBUDDY, 0, 0)
            if is_valid_hwnd(buddy):
                bl = win32gui.GetWindowRect(buddy)[0]
                try_add(buddy, bl, r["top"])

    if not cands:
        return None
    cands.sort(key=lambda x: x[0])
    return cands[0][1]


def _computed_n2_click_point(n1_hw, d1_cb, plus_anchor):
    """由 n1/φ1/「+」几何推算 n2 Spinner 中心（必须在 φ1 右缘之外）。"""
    n1l, n1t, n1r, n1b = win32gui.GetWindowRect(n1_hw)
    _, _, d1r, _ = win32gui.GetWindowRect(d1_cb)
    pl = plus_anchor["left"]
    n1_w = max(n1r - n1l, 24)
    x = max(d1r + 16, pl + 14) + n1_w // 2
    y = (n1t + n1b) // 2
    return x, y


def _fill_edit_by_screen_click(
    main_hw, x, y, value, label: str, d1_cb=None,
    n1_hw=None, plus_anchor=None, results_log_hw=None,
) -> bool:
    """物理点击 + 键盘输入（不 Tab）；校验焦点不是 Results 大文本框。"""
    print(f"  → {label} 点击输入 ({int(x)}, {int(y)})")
    safe_click(main_hw, int(x), int(y))
    time.sleep(SLEEP_STIRRUP_FIELD)
    focus = get_expert_focus(main_hw)
    _log_hwnd(f"{label} 点击焦点", focus)
    if not _is_valid_n2_hwnd(
        focus, n1_hw, d1_cb, plus_anchor, results_log_hw,
    ):
        print(f"  ✗ {label} 点击落点不是箍筋 Spinner（可能点到 Results 输出框）")
        return False
    send_ctrl_a()
    send_text(str(int(value)))
    time.sleep(0.08)
    focus = get_expert_focus(main_hw)
    if is_valid_hwnd(focus) and win32gui.GetClassName(focus) in EDIT_CLASSES:
        _notify_edit_change(focus)
    got = _read_edit_text(focus) if is_valid_hwnd(focus) else ""
    try:
        ok = int(float(str(got).replace(",", "."))) == int(value)
    except (TypeError, ValueError):
        ok = str(got).strip() == str(int(value))
    if ok and results_log_hw and is_valid_hwnd(results_log_hw):
        log_got = _read_edit_text(results_log_hw)
        if str(int(value)) in log_got and len(log_got) > 12:
            print(f"  ✗ {label} 误写入 Results 输出框")
            return False
    if ok:
        print(f"  → {label}={int(value)} ✓")
    else:
        print(f"  ✗ {label} 期望 {int(value)}，读回 {got!r}")
    if ok and d1_cb and label == "n2":
        d1_got = _read_combo_text(d1_cb)
        bare = re.sub(r"^[φΦøØ]\s*", "", d1_got).strip()
        if bare == str(int(value)):
            print(f"  ✗ n2={int(value)} 误写入 φ1（φ1 现为 {d1_got!r}）")
            return False
    return ok


def _is_valid_d2_hwnd(d2_hw, d1_cb) -> bool:
    if not is_valid_hwnd(d2_hw):
        return False
    if d2_hw == d1_cb:
        return False
    cls = win32gui.GetClassName(d2_hw)
    if cls == COMBOBOX_CLASS:
        return True
    if cls in EDIT_CLASSES:
        parent = win32gui.GetParent(d2_hw)
        if parent and win32gui.GetClassName(parent) == COMBOBOX_CLASS:
            return parent != d1_cb
    return False


def _find_plus_static(rows, grid_rect):
    gl, _, gr, gb = grid_rect
    mid_x = (gl + gr) // 2
    for r in rows:
        if r["class"] != "Static":
            continue
        if r["top"] < gb - 5 or r["left"] > mid_x:
            continue
        if (r["text"] or "").strip() == "+":
            return {"left": r["left"], "top": r["top"]}
    return None


def _find_d2_combo_deep(rows, plus_anchor, d1_cb, n2_hw, enabled_only=False):
    if not plus_anchor or not is_valid_hwnd(n2_hw):
        return None
    pt = plus_anchor["top"]
    n2l, _, n2r, _ = win32gui.GetWindowRect(n2_hw)
    min_x = (n2l + n2r) // 2 + 4
    cands = []
    for r in rows:
        if r["class"] != COMBOBOX_CLASS or r["hwnd"] == d1_cb:
            continue
        if abs(r["top"] - pt) > 22:
            continue
        cx = r["left"] + r["width"] // 2
        if cx <= min_x:
            continue
        if enabled_only and not win32gui.IsWindowEnabled(r["hwnd"]):
            continue
        cands.append((cx, r["hwnd"]))
    if not cands:
        return None
    cands.sort()
    bar = [(cx, h) for cx, h in cands if _is_rebar_diameter_combo(h)]
    pick = bar or cands
    return pick[0][1]


def _computed_d2_click_point(n2_hw, plus_anchor):
    n2l, n2t, n2r, n2b = win32gui.GetWindowRect(n2_hw)
    n2_w = max(n2r - n2l, 24)
    x = n2r + 28
    y = (n2t + n2b) // 2
    return x, y


def _resolve_n2_hwnd(main_hw, rows, n1_hw, d1_cb, plus_anchor=None, results_log_hw=None):
    """n2 定位：深度扫描 → 几何点击推算（排除 Results 输出框）。"""
    n2_hw = _find_n2_edit_deep(
        rows, n1_hw, d1_cb, plus_anchor, results_log_hw,
    )
    if n2_hw:
        print("  → n2 深度扫描定位")
        _log_hwnd("n2 扫描", n2_hw)
        return n2_hw

    _log_stirrup_row_candidates(
        rows, n1_hw, plus_anchor, "n2 行", d1_cb, results_log_hw,
    )

    if plus_anchor:
        x, y = _computed_n2_click_point(n1_hw, d1_cb, plus_anchor)
        hw = _hwnd_after_screen_click(main_hw, x, y)
        ed = _edit_from_spin_hwnd(hw) or hw
        if _is_valid_n2_hwnd(ed, n1_hw, d1_cb, plus_anchor, results_log_hw):
            _log_hwnd("n2 推算点击焦点", ed)
            return ed

    if has_stirrup_coords():
        x, y = stirrup_screen_point(main_hw, "N2")
        hw = _hwnd_after_screen_click(main_hw, x, y)
        ed = _edit_from_spin_hwnd(hw) or hw
        if _is_valid_n2_hwnd(ed, n1_hw, d1_cb, plus_anchor, results_log_hw):
            print(f"  → n2 校准坐标 ({int(x)}, {int(y)})")
            _log_hwnd("n2 校准点击焦点", ed)
            return ed
        print("  → 校准坐标未命中有效 n2（可能点到 Results 输出框）")

    return None


def _resolve_d2_hwnd(main_hw, rows, n2_hw, d1_cb, plus_anchor=None):
    """φ2 定位：深度扫描 → 推算坐标 → 校准坐标（不用 Tab）。"""
    _click_hwnd(main_hw, n2_hw, combo=False)
    _notify_edit_change(n2_hw)
    time.sleep(SLEEP_STIRRUP_N2)

    deadline = time.time() + 2.5
    while time.time() < deadline:
        rows = _get_design_rows_deep(main_hw)
        d2 = _find_d2_combo_deep(rows, plus_anchor, d1_cb, n2_hw, enabled_only=True)
        if d2:
            return d2
        time.sleep(0.1)

    if plus_anchor:
        x, y = _computed_d2_click_point(n2_hw, plus_anchor)
        hw = _hwnd_after_screen_click(main_hw, x, y)
        if _is_valid_d2_hwnd(hw, d1_cb):
            return hw

    if has_stirrup_coords():
        x, y = stirrup_screen_point(main_hw, "D2")
        hw = _hwnd_after_screen_click(main_hw, x, y)
        if _is_valid_d2_hwnd(hw, d1_cb) and win32gui.IsWindowEnabled(hw):
            return hw

    print("  ✗ φ2 下拉框未启用或无法定位")
    return None


def _force_edit_focus(main_hw, hwnd):
    try:
        win32gui.SetForegroundWindow(main_hw)
        win32gui.SetFocus(hwnd)
    except Exception:
        pass
    time.sleep(0.05)


def _edit_int_matches(hwnd, value) -> bool:
    got = _read_edit_text(hwnd)
    try:
        return int(float(str(got).replace(",", "."))) == int(value)
    except (TypeError, ValueError):
        return str(got).strip() == str(int(value))


def _send_key_to_hwnd(hwnd, vk):
    """向指定控件发送按键（与荷载表 grid_send_key 同理，不用全局 SendInput）。"""
    if not is_valid_hwnd(hwnd):
        send_vk(vk)
        return
    scan = win32api.MapVirtualKey(vk, 0)
    lp_down = 1 | (scan << 16)
    lp_up = lp_down | (1 << 30) | (1 << 31)
    win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk, lp_down)
    time.sleep(0.03)
    win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk, lp_up)
    time.sleep(0.05)


def _human_commit_edit(main_hw, hwnd):
    """人类操作：Enter 确认当前输入框，退出编辑态以便 Tab 切到下一控件。"""
    focus = get_expert_focus(main_hw) or hwnd
    _send_key_to_hwnd(focus, win32con.VK_RETURN)
    time.sleep(0.1)


def _human_type_into_hwnd(main_hw, hwnd, value, label: str) -> bool:
    """人类操作：鼠标点击 → 全选 → 键盘输入 → Enter 确认。"""
    x, y = _hwnd_click_point(hwnd, combo=False)
    safe_click(main_hw, x, y)
    time.sleep(SLEEP_STIRRUP_FIELD)
    try:
        win32gui.SetForegroundWindow(main_hw)
        win32gui.SetFocus(hwnd)
    except Exception:
        pass
    time.sleep(0.05)
    send_ctrl_a()
    send_text(str(int(value)))
    time.sleep(0.08)
    _human_commit_edit(main_hw, hwnd)
    if _edit_int_matches(hwnd, value):
        print(f"  → {label}={int(value)} ✓ (点击+键盘+Enter)")
        return True
    print(f"  ✗ {label} 期望 {int(value)}，读回 {_read_edit_text(hwnd)!r}")
    return False


def _human_tab(main_hw, steps: int) -> Optional[int]:
    """
    人类 Tab：向当前焦点控件 PostMessage Tab；无效则 WM_NEXTDLGCTL 切到下一控件。
    （全局 SendInput Tab 到不了 Expert 对话框内部，这是与手敲键盘的核心区别。）
    """
    design_hw = None
    try:
        design_hw = get_design_page_hwnd(main_hw)
    except RuntimeError:
        design_hw = main_hw

    before = get_expert_focus(main_hw)
    for step in range(1, steps + 1):
        focus = get_expert_focus(main_hw) or before
        if is_valid_hwnd(focus):
            _send_key_to_hwnd(focus, win32con.VK_TAB)
            time.sleep(0.12)
            after = get_expert_focus(main_hw)
            if after and after != focus:
                _log_hwnd(f"Tab×{step} 焦点", after)
                before = after
                continue
        if design_hw:
            win32gui.SendMessage(design_hw, WM_NEXTDLGCTL, 1, 0)
            time.sleep(0.12)
            after = get_expert_focus(main_hw)
            _log_hwnd(f"DlgTab×{step} 焦点", after)
            if after:
                before = after
    return get_expert_focus(main_hw)


def _set_spin_edit_value(main_hw, hwnd, value, label: str) -> bool:
    """Spinner 伙伴 Edit：点击聚焦 → WM_SETTEXT → 键盘回退。"""
    if not is_valid_hwnd(hwnd):
        return False
    _click_hwnd(main_hw, hwnd, combo=False)
    try:
        win32gui.SendMessage(hwnd, win32con.EM_SETSEL, 0, -1)
    except Exception:
        pass
    set_edit_text(hwnd, int(value))
    _notify_edit_change(hwnd)
    time.sleep(0.12)
    if _edit_int_matches(hwnd, value):
        print(f"  → {label}={int(value)} ✓ (WM_SETTEXT)")
        return True
    return _human_type_into_hwnd(main_hw, hwnd, value, label)


def _fill_n2_value(
    main_hw, hwnd, value, n1_hw, d1_cb, blocked, plus_anchor,
) -> bool:
    """写入 n2：句柄直写 → 屏幕点击 → 校准坐标。"""
    if _set_spin_edit_value(main_hw, hwnd, value, "n2"):
        return True
    if _stirrup_set_edit(
        main_hw, hwnd, value, "n2", d1_cb=d1_cb, results_log_hw=blocked,
    ):
        return True
    x, y = _hwnd_click_point(hwnd, combo=False)
    if _fill_edit_by_screen_click(
        main_hw, x, y, value, "n2",
        d1_cb=d1_cb, n1_hw=n1_hw, plus_anchor=plus_anchor,
        results_log_hw=blocked,
    ):
        return True
    if has_stirrup_coords():
        x, y = stirrup_screen_point(main_hw, "N2")
        return _fill_edit_by_screen_click(
            main_hw, x, y, value, "n2",
            d1_cb=d1_cb, n1_hw=n1_hw, plus_anchor=plus_anchor,
            results_log_hw=blocked,
        )
    return False


def _is_tab_focus_n2_candidate(focus, n1_hw, d1_cb, plus_anchor, blocked) -> bool:
    """Tab 后焦点是否可能是 n2（同行 Edit、在 n1 右侧、非 Results/φ1 内部）。"""
    if not is_valid_hwnd(focus) or focus == n1_hw or focus == blocked:
        return False
    if win32gui.GetClassName(focus) not in EDIT_CLASSES:
        return False
    if win32gui.GetClassName(focus) in RICHEDIT_CLASSES:
        return False
    if d1_cb and _is_d1_combo_edit(focus, d1_cb):
        return False
    if d1_cb and focus in _combo_inner_edit_hwnds(d1_cb):
        return False
    if d1_cb and _hwnd_is_descendant_of(focus, d1_cb):
        fl, _, fr, _ = win32gui.GetWindowRect(focus)
        d1l, _, d1r, _ = win32gui.GetWindowRect(d1_cb)
        if (fl + fr) // 2 < d1r - 4:
            return False
    return _is_valid_n2_hwnd(
        focus, n1_hw, d1_cb, plus_anchor, blocked,
    )


def _is_tab_focus_likely_n2(focus, n1_hw, d1_cb, blocked) -> bool:
    """Tab 离开 n1 后宽松识别：同行右侧小编辑框即尝试写入（避免几何阈值误杀）。"""
    if not is_valid_hwnd(focus) or focus == n1_hw or focus == blocked:
        return False
    cls = win32gui.GetClassName(focus)
    if cls in RICHEDIT_CLASSES or cls not in EDIT_CLASSES:
        return False
    l, t, r, b = win32gui.GetWindowRect(focus)
    w, h = r - l, b - t
    if w > 150 or h > 50:
        return False
    if d1_cb and _is_d1_combo_edit(focus, d1_cb):
        return False
    if d1_cb and focus in _combo_inner_edit_hwnds(d1_cb):
        return False
    if not is_valid_hwnd(n1_hw):
        return True
    n1l, n1t, n1r, n1b = win32gui.GetWindowRect(n1_hw)
    if abs(t - n1t) > 16:
        return False
    return (l + r) // 2 > (n1l + n1r) // 2 + 4


def _tab_find_and_fill_n2(
    main_hw, n1_hw, d1_cb, blocked, plus_anchor, n2_input, max_tabs,
) -> Optional[int]:
    """逐次 Tab，落在 n2 候选 Edit 上则写入。"""
    before = get_expert_focus(main_hw)
    for step in range(1, max_tabs + 1):
        focus = _human_tab(main_hw, 1)
        if focus == before:
            continue
        strict = _is_tab_focus_n2_candidate(
            focus, n1_hw, d1_cb, plus_anchor, blocked,
        )
        loose = _is_tab_focus_likely_n2(focus, n1_hw, d1_cb, blocked)
        if not strict and not loose:
            continue
        tag = "候选 n2" if strict else "宽松 n2"
        _log_hwnd(f"Tab×{step} {tag}", focus)
        if _fill_n2_value(
            main_hw, focus, n2_input, n1_hw, d1_cb, blocked, plus_anchor,
        ):
            return focus
        before = focus
    return None


def _fill_n1_n2_stirrup(
    main_hw, ctrls, n1, n2_input, tab_to_n2, blocked, rows,
) -> bool:
    """n1 键盘输入 → Tab 找 n2 → WM_SETTEXT；失败则深度扫描。"""
    n1_hw = ctrls["n1_hw"]
    d1_cb = ctrls["d1_cb"]
    plus_anchor = ctrls.get("stirrup_plus")
    if not plus_anchor:
        plus_anchor = _find_plus_static(rows, ctrls["grid_rect"])
        ctrls["stirrup_plus"] = plus_anchor

    if not _human_type_into_hwnd(main_hw, n1_hw, n1, "n1"):
        return False

    n2_hw = ctrls.get("n2_hw")
    if is_valid_hwnd(n2_hw) and not _is_valid_n2_hwnd(
        n2_hw, n1_hw, d1_cb, plus_anchor, blocked,
    ):
        _log_hwnd("扫描 n2 无效(φ1内Edit?)", n2_hw)
        n2_hw = None
    elif is_valid_hwnd(n2_hw) and _is_valid_n2_hwnd(
        n2_hw, n1_hw, d1_cb, plus_anchor, blocked,
    ):
        _log_hwnd("扫描 n2", n2_hw)
        if not _fill_n2_value(
            main_hw, n2_hw, n2_input, n1_hw, d1_cb, blocked, plus_anchor,
        ):
            n2_hw = None

    if not n2_hw:
        print(f"  → n1→Tab 找 n2（最多 {tab_to_n2} 次）")
        n2_hw = _tab_find_and_fill_n2(
            main_hw, n1_hw, d1_cb, blocked, plus_anchor,
            n2_input, tab_to_n2,
        )

    if not n2_hw:
        print("  → Tab 未写入 n2，深度扫描 Spinner…")
        n2_hw = _find_n2_edit_deep(
            rows, n1_hw, d1_cb, plus_anchor, blocked,
        )
        if n2_hw:
            _log_hwnd("深度扫描 n2", n2_hw)
            if not _fill_n2_value(
                main_hw, n2_hw, n2_input, n1_hw, d1_cb, blocked, plus_anchor,
            ):
                n2_hw = None
        if not n2_hw:
            n2_hw = _resolve_n2_hwnd(
                main_hw, rows, n1_hw, d1_cb, plus_anchor, blocked,
            )
            if n2_hw:
                _log_hwnd("推算 n2", n2_hw)
                if not _fill_n2_value(
                    main_hw, n2_hw, n2_input, n1_hw, d1_cb, blocked, plus_anchor,
                ):
                    n2_hw = None

    if not n2_hw:
        print("  ✗ 未填入 n2")
        return False

    d1_got = _read_combo_text(d1_cb)
    bare = re.sub(r"^[φΦøØ]\s*", "", d1_got).strip()
    if int(n2_input) > 0 and bare == str(int(n2_input)):
        print(f"  ✗ n2 误写入 φ1（φ1 现为 {d1_got!r}）")
        return False

    ctrls["n2_hw"] = n2_hw
    return True


def fill_stirrup_sequential(main_hw, ctrls, n1, d_bar1):
    """箍筋填写：仅 n1 + φ1（WM_SETTEXT + ComboBox，不操作 n2/φ2）。"""
    main_hw = ensure_expert_visible(main_hw)
    if not refresh_stirrup_hwnds(main_hw, ctrls):
        print("  ✗ 未找到箍筋 n1/φ1 控件")
        return False

    _log_stirrup_hwnds(ctrls)
    n1_hw = ctrls["n1_hw"]
    d1_cb = ctrls["d1_cb"]

    blocked = ctrls.get("results_log_hwnd")
    if is_valid_hwnd(blocked):
        win32gui.EnableWindow(blocked, False)
        time.sleep(0.08)

    try:
        with expert_keyboard_session(main_hw):
            if not _stirrup_set_edit(
                main_hw, n1_hw, n1, "n1", results_log_hw=blocked,
            ):
                return False
            time.sleep(0.08)
            if not _stirrup_set_combo(main_hw, d1_cb, d_bar1, "φ1"):
                return False
            print("  → 箍筋输入完成")
            return True
    finally:
        if is_valid_hwnd(blocked):
            win32gui.EnableWindow(blocked, True)
    return False


def collect_controls(parent):
    rows = []

    def cb(hw, _):
        left, top, right, bottom = win32gui.GetWindowRect(hw)
        rows.append({
            "hwnd": hw,
            "class": win32gui.GetClassName(hw),
            "text": win32gui.GetWindowText(hw),
            "ctrl_id": win32gui.GetDlgCtrlID(hw),
            "top": top, "left": left,
            "width": right - left, "height": bottom - top,
        })
        return True

    win32gui.EnumChildWindows(parent, cb, None)
    rows.sort(key=lambda r: (r["top"], r["left"]))
    return rows


def collect_all_controls(main_hw):
    """递归收集主窗口下所有控件（Design 子窗口消失时的后备方案）"""
    rows = []

    def walk(hwnd):
        try:
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            rows.append({
                "hwnd": hwnd,
                "class": win32gui.GetClassName(hwnd),
                "text": win32gui.GetWindowText(hwnd),
                "ctrl_id": win32gui.GetDlgCtrlID(hwnd),
                "top": top, "left": left,
                "width": right - left, "height": bottom - top,
            })
            win32gui.EnumChildWindows(hwnd, lambda c, _: walk(c) or True, None)
        except Exception:
            pass

    walk(main_hw)
    rows.sort(key=lambda r: (r["top"], r["left"]))
    return rows


def get_page_rows(main_hw):
    """优先 Design 页；找不到则扫描整个 Expert 主窗口"""
    try:
        design_hw = get_design_page_hwnd(main_hw)
        return collect_controls(design_hw), design_hw
    except RuntimeError:
        return collect_all_controls(main_hw), main_hw


def parse_load_grid_rect(rows):
    gxwnds = [r for r in rows if r["class"] == "GXWND" and r["width"] > 50]
    if gxwnds:
        r = max(gxwnds, key=lambda x: x["width"] * x["height"])
        rect = (r["left"], r["top"], r["left"] + r["width"], r["top"] + r["height"])
        return rect, r["hwnd"]
    for r in rows:
        if r["class"] in BUTTON_CLASS and "loads" in r["text"].lower() and "kn" in r["text"].lower():
            rect = (
                r["left"] + 95, r["top"] + 23,
                r["left"] + r["width"] - 15, r["top"] + r["height"] - 77,
            )
            return rect, None
    raise RuntimeError("未找到荷载表，请确认在 Design 页")


def find_button_hwnd(rows, *keywords):
    for r in rows:
        if r["class"] not in BUTTON_CLASS:
            continue
        text = r["text"].lower().replace("&", "")
        if any(k.lower() in text for k in keywords):
            return r["hwnd"]
    return None


def find_edit_by_ctrl_id(rows, ctrl_id):
    for r in rows:
        if r["ctrl_id"] == ctrl_id and r["class"] in EDIT_CLASSES:
            return r["hwnd"]
    return None


def find_dimension_edits(rows, design_hw):
    by_id = {k: find_edit_by_ctrl_id(rows, v) for k, v in DIM_CTRL_IDS.items()}
    if all(by_id.values()):
        return by_id["edit_b"], by_id["edit_h"], by_id["edit_d1"], by_id["edit_d2"]

    panel = None
    for r in rows:
        if r["class"] in BUTTON_CLASS and "dimensions" in r["text"].lower():
            panel = (r["left"], r["top"], r["left"] + r["width"], r["top"] + r["height"])
            break

    if panel:
        pl, pt, pr, pb = panel
        in_panel = [r for r in rows if r["class"] in EDIT_CLASSES and r["width"] > 10
                    and pl <= r["left"] <= pr and pt <= r["top"] <= pb]
    else:
        dl, dt, dr, db = win32gui.GetWindowRect(design_hw)
        in_panel = [r for r in rows if r["class"] in EDIT_CLASSES and r["width"] > 10
                    and r["left"] >= dl + (dr - dl) // 2]

    if len(in_panel) < 4:
        raise RuntimeError("未找到 b/h/d1/d2，请运行 dump_controls.py")

    in_panel.sort(key=lambda r: (r["top"], r["left"]))
    left_x = min(r["left"] for r in in_panel)
    left_col = sorted([r for r in in_panel if r["left"] <= left_x + 80], key=lambda r: r["top"])
    bottom = sorted([r for r in in_panel if r["top"] >= max(r["top"] for r in in_panel) - 8],
                    key=lambda r: r["left"])
    return left_col[0]["hwnd"], left_col[1]["hwnd"], bottom[0]["hwnd"], bottom[1]["hwnd"]


def find_results_log_edit(rows, grid_rect):
    """Results 区大文本框（会抢走键盘焦点，填荷载时需临时禁用）"""
    _, _, _, gb = grid_rect
    best_hw, best_area = None, 0
    for r in rows:
        if r["class"] not in EDIT_CLASSES and r["class"] not in RICHEDIT_CLASSES:
            continue
        if r["top"] < gb - 5:
            continue
        if r["width"] < 120 or r["height"] < 25:
            continue
        area = r["width"] * r["height"]
        if area > best_area:
            best_area = area
            best_hw = r["hwnd"]
    return best_hw


def find_results_as_edits(rows, grid_rect):
    """Results 区 As1、As2 数值框（荷载表下方左侧，非右侧 d1/d2）"""
    gl, _, gr, gb = grid_rect
    mid_x = (gl + gr) // 2
    dim_ids = set(DIM_CTRL_IDS.values())
    candidates = []
    for r in rows:
        if r["class"] not in EDIT_CLASSES:
            continue
        if r["top"] < gb - 5 or r["left"] > mid_x:
            continue
        if r["ctrl_id"] in dim_ids:
            continue
        if not (25 <= r["width"] <= 100 and 12 <= r["height"] <= 35):
            continue
        candidates.append(r)
    candidates.sort(key=lambda r: (r["top"], r["left"]))
    if len(candidates) >= 2:
        return candidates[0]["hwnd"], candidates[1]["hwnd"]
    return None, None


def _read_edit_value(hwnd):
    if not is_valid_hwnd(hwnd):
        return None
    text = win32gui.GetWindowText(hwnd).strip()
    if not text:
        return None
    try:
        return float(text.replace(",", "."))
    except ValueError:
        return text


def _parse_reinforcement_statics(rows, grid_rect):
    """从 Results 区 Static 文本解析 ρ、ρmin、ρmax（%）"""
    _, _, _, gb = grid_rect
    rho, rho_min, rho_max = None, None, None
    for r in rows:
        if r["class"] != "Static" or r["top"] < gb - 5:
            continue
        t = (r["text"] or "").replace(",", ".")
        if not t:
            continue
        tl = t.lower()
        if "reinforcement ratio" in tl and "min" not in tl and "max" not in tl:
            m = re.search(r"=\s*([\d.]+)\s*%", t)
            if m:
                rho = float(m.group(1))
        elif "minimum" in tl and "reinforcement" in tl:
            m = re.search(r"([\d.]+)\s*%", t)
            if m:
                rho_min = float(m.group(1))
        elif "maximum" in tl and "reinforcement" in tl:
            m = re.search(r"([\d.]+)\s*%", t)
            if m:
                rho_max = float(m.group(1))
    return rho, rho_min, rho_max


def _read_edit_text(hwnd) -> str:
    if not is_valid_hwnd(hwnd):
        return ""
    length = win32gui.SendMessage(hwnd, win32con.WM_GETTEXTLENGTH, 0, 0)
    if length <= 0:
        return win32gui.GetWindowText(hwnd).strip()
    buf = ctypes.create_unicode_buffer(length + 2)
    win32gui.SendMessage(hwnd, win32con.WM_GETTEXT, length + 1, buf)
    return buf.value.strip()


def _find_edit_right_of_static(rows, static_row, x_limit):
    st_top, st_left = static_row["top"], static_row["left"]
    best, best_dx = None, 99999
    for r in rows:
        if r["class"] not in EDIT_CLASSES:
            continue
        if abs(r["top"] - st_top) > 14:
            continue
        if r["left"] <= st_left + 5:
            continue
        if r["left"] > x_limit:
            continue
        dx = r["left"] - st_left
        if dx < best_dx:
            best_dx = dx
            best = r
    return best["hwnd"] if best else None


def _parse_shear_verdict(log_text: str) -> str:
    if not log_text:
        return "已计算"
    t = log_text.lower()
    if any(k in t for k in ("not fulfilled", "non ok", "failed", "incorrect", "未满足")):
        return "NON OK"
    if any(k in t for k in ("condition is fulfilled", " ok", "满足", "verified")):
        return "OK"
    return "已计算"


def _short_log_summary(log_text: str, max_len: int = 150) -> str:
    if not log_text:
        return ""
    for line in log_text.replace("\r", "").split("\n"):
        s = line.strip()
        if s:
            return s[:max_len]
    return ""


def read_shear_expert_results(main_hw, grid_rect, log_hwnd=None) -> dict:
    """Calculate 后读取 St、St,max 及 Results 区摘要。"""
    rows, _ = get_page_rows(main_hw)
    gl, _, gr, gb = grid_rect
    x_limit = (gl + gr) // 2 + 120
    st_hw, st_max_hw = None, None
    for r in rows:
        if r["class"] != "Static" or r["top"] < gb - 5:
            continue
        t = (r["text"] or "").strip()
        tl = re.sub(r"\s+", "", t.lower())
        if "st,max" in tl or "stmax" in tl:
            st_max_hw = _find_edit_right_of_static(rows, r, x_limit)
        elif tl.startswith("st") and "max" not in tl:
            st_hw = _find_edit_right_of_static(rows, r, x_limit)

    st_val = _read_edit_value(st_hw)
    st_max_val = _read_edit_value(st_max_hw)
    log_text = _read_edit_text(log_hwnd) if log_hwnd else ""
    verdict = _parse_shear_verdict(log_text)
    summary = _short_log_summary(log_text)

    out = {
        "St_cm": st_val,
        "St_max_cm": st_max_val,
        "计算状态": verdict,
        "结果摘要": summary,
    }
    print(
        f"  → 抗剪结果: St={st_val}  St,max={st_max_val}  "
        f"状态={verdict}"
    )
    if summary:
        print(f"  → 摘要: {summary[:80]}...")
    return out


def summarize_loads(load_rows: List[dict]) -> tuple:
    """返回 (组数, 类型列表, 控制剪力 V, 对应 N)。"""
    n = len(load_rows)
    if not load_rows:
        return n, "", None, None
    types = ", ".join(str(lr.get("load_type", "")) for lr in load_rows)
    ctrl = max(load_rows, key=lambda x: abs(_f(x.get("V_kN"), 0)))
    return n, types, ctrl.get("V_kN"), ctrl.get("N_kN")


def _f(v, default=0.0):
    try:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return default
        return float(v)
    except (TypeError, ValueError):
        return default


def build_result_summary_row(case: dict, expert_out: dict, rtf_path: str) -> dict:
    n_load, load_types, v_ctrl, n_ctrl = summarize_loads(case["loads"])
    row = {
        "区域": case["region"],
        "编号": case["code"],
        "b_cm": case["b_cm"],
        "h_cm": case["h_cm"],
        "d_cm": case["d1_cm"],
        "n1": case["n1"],
        "d_bar1": case["d_bar1"],
        "n2": case["n2"],
        "d_bar2": case["d_bar2"],
        "荷载组数": n_load,
        "荷载类型": load_types,
        "V_control_kN": round(v_ctrl, 2) if v_ctrl is not None else "",
        "N_control_kN": round(n_ctrl, 2) if n_ctrl is not None else "",
        "St_cm": expert_out.get("St_cm", ""),
        "St_max_cm": expert_out.get("St_max_cm", ""),
        "计算状态": expert_out.get("计算状态", ""),
        "结果摘要": expert_out.get("结果摘要", ""),
        "RTF路径": rtf_path or "",
    }
    return row


def export_result_summary(results: List[dict], path: str):
    """写入结果汇总 Excel（单 sheet）。"""
    df = pd.DataFrame(results)
    for c in RESULT_SUMMARY_COLUMNS:
        if c not in df.columns:
            df[c] = ""
    df = df[RESULT_SUMMARY_COLUMNS]
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="结果汇总", index=False)
    print(f"结果汇总 → {path}（{len(df)} 行）")


def find_checkbox_by_keyword(rows, keyword, exclude=()):
    kw = keyword.lower()
    ex = tuple(x.lower() for x in exclude)
    for r in rows:
        if r["class"] not in BUTTON_CLASS:
            continue
        t = r["text"].lower().replace("&", "")
        if kw not in t:
            continue
        if any(e in t for e in ex):
            continue
        return r["hwnd"]
    return None


def find_stirrup_controls(rows, grid_rect):
    """Results 区箍筋：n1、d_bar1、n2、d_bar2（以「+」为锚点定位同一行）。"""
    gl, _, gr, gb = grid_rect
    mid_x = (gl + gr) // 2
    dim_ids = set(DIM_CTRL_IDS.values())

    plus_row = None
    for r in rows:
        if r["class"] != "Static":
            continue
        if r["top"] < gb - 5 or r["left"] > mid_x:
            continue
        if (r["text"] or "").strip() != "+":
            continue
        pt, pl = r["top"], r["left"]
        row_items = []
        for rr in rows:
            if abs(rr["top"] - pt) > 14:
                continue
            if rr["ctrl_id"] in dim_ids:
                continue
            if rr["class"] in EDIT_CLASSES and 12 <= rr["width"] <= 80:
                row_items.append(("e", rr))
            elif rr["class"] == COMBOBOX_CLASS and 28 <= rr["width"] <= 100:
                row_items.append(("c", rr))
        left_e = [x for k, x in row_items if k == "e" and x["left"] < pl - 5]
        left_c = [x for k, x in row_items if k == "c" and x["left"] < pl - 5]
        if left_e and left_c:
            plus_row = r
            break

    if plus_row:
        pt, pl = plus_row["top"], plus_row["left"]
        row_items = []
        for r in rows:
            if abs(r["top"] - pt) > 18:
                continue
            if r["ctrl_id"] in dim_ids:
                continue
            if r["class"] in EDIT_CLASSES and 12 <= r["width"] <= 80:
                row_items.append(("e", r))
            elif r["class"] == COMBOBOX_CLASS and 20 <= r["width"] <= 120:
                row_items.append(("c", r))

        edits_raw = sorted([r for k, r in row_items if k == "e"], key=lambda x: x["left"])
        combos = sorted([r for k, r in row_items if k == "c"], key=lambda x: x["left"])
        seen_e = {r["hwnd"] for r in edits_raw}
        seen_c = {c["hwnd"] for c in combos}
        for r in rows:
            if abs(r["top"] - pt) > 24:
                continue
            if r["ctrl_id"] in dim_ids:
                continue
            if r["class"] in EDIT_CLASSES and 12 <= r["width"] <= 90:
                if r["hwnd"] not in seen_e:
                    edits_raw.append(r)
                    seen_e.add(r["hwnd"])
            elif r["class"] == COMBOBOX_CLASS and 20 <= r["width"] <= 120:
                if r["hwnd"] not in seen_c:
                    combos.append(r)
                    seen_c.add(r["hwnd"])
        edits_raw.sort(key=lambda x: x["left"])
        combos.sort(key=lambda x: x["left"])
        edits = _filter_combo_inner_edits(edits_raw, combos)

        edits_left = [e for e in edits if e["left"] < pl - 5]
        combos_left = [c for c in combos if c["left"] < pl - 5]
        d1_row = None
        bar_left = [c for c in combos_left if _is_rebar_diameter_combo(c["hwnd"])]
        pick_left = bar_left or combos_left
        if pick_left:
            d1_row = pick_left[0]
        d1 = d1_row["hwnd"] if d1_row else None

        n1 = edits_left[0]["hwnd"] if edits_left else None
        n2 = _resolve_n2_edit(edits, pl, d1_row, combos)
        if n2 and d1 and _is_d1_combo_edit(n2, d1):
            n2 = None
        if not n2 and d1_row:
            inner = set()
            for c in combos:
                inner |= _combo_inner_edit_hwnds(c["hwnd"])
            min_x = _n2_min_center_x(pl, d1_row)
            for r in rows:
                if r["class"] not in EDIT_CLASSES:
                    continue
                if r["top"] < gb - 5 or abs(r["top"] - pt) > 24:
                    continue
                if n1 and r["hwnd"] == n1:
                    continue
                if r["hwnd"] in inner:
                    continue
                if any(_hwnd_is_descendant_of(r["hwnd"], c["hwnd"]) for c in combos):
                    continue
                if _row_center_x(r) <= min_x:
                    continue
                if any(_edit_inside_combo_box(r, c) for c in combos):
                    continue
                if d1 and _is_d1_combo_edit(r["hwnd"], d1):
                    continue
                n2 = r["hwnd"]
                break
        d2 = _resolve_d2_combo(combos, pl, d1_row, pt, rows)

        if n1 and d1:
            return n1, d1, n2, d2, {"left": pl, "top": pt}

    edits, combos = [], []
    for r in rows:
        if r["top"] < gb - 5 or r["left"] > mid_x:
            continue
        if r["ctrl_id"] in dim_ids:
            continue
        if r["class"] in EDIT_CLASSES and 12 <= r["width"] <= 80 and 12 <= r["height"] <= 40:
            edits.append(r)
        elif r["class"] == COMBOBOX_CLASS and 28 <= r["width"] <= 100:
            combos.append(r)
    if not edits and not combos:
        return None, None, None, None, None
    items = sorted(
        [("e", r) for r in edits] + [("c", r) for r in combos],
        key=lambda x: (x[1]["top"], x[1]["left"]),
    )
    bands: Dict[int, list] = defaultdict(list)
    for kind, r in items:
        bands[r["top"] // 12].append((kind, r))
    best = max(bands.values(), key=lambda lst: len(lst))
    best.sort(key=lambda x: x[1]["left"])
    e_list = [r for k, r in best if k == "e"]
    c_list = [r for k, r in best if k == "c"]
    n1 = e_list[0]["hwnd"] if e_list else None
    d1 = c_list[0]["hwnd"] if c_list else None
    n2 = e_list[1]["hwnd"] if len(e_list) > 1 else None
    d2 = c_list[1]["hwnd"] if len(c_list) > 1 else None
    return n1, d1, n2, d2, None


def build_control_map(main_hw):
    main_hw = ensure_expert_visible(main_hw)
    rows, page_hw = get_page_rows(main_hw)
    grid_rect, grid_hwnd = parse_load_grid_rect(rows)
    edit_b, edit_h, edit_d1, edit_d2 = find_dimension_edits(rows, page_hw)
    edit_d = edit_d1  # 抗剪界面仅一个 d 输入框

    btn_calc = find_button_hwnd(rows, "calculate", "c a l c u l a t e")
    btn_note = find_button_hwnd(rows, "note", "n o t e")
    chk_seismic = find_button_hwnd(rows, "seismic")
    chk_shear = find_checkbox_by_keyword(rows, "shear", exclude=("torsion",))
    chk_torsion = find_checkbox_by_keyword(rows, "torsion")
    radio_critical = find_checkbox_by_keyword(rows, "critical", exclude=("non",))
    radio_column = find_checkbox_by_keyword(rows, "column", exclude=("beam",))
    n1_hw, d1_cb, n2_hw, d2_cb, stirrup_plus = find_stirrup_controls(rows, grid_rect)

    missing = [n for n, h in [("Calculate", btn_calc), ("Note", btn_note)] if not h]
    if missing:
        raise RuntimeError(f"未找到按钮: {missing}")

    return {
        "edit_b": edit_b, "edit_h": edit_h, "edit_d": edit_d,
        "chk_seismic": chk_seismic,
        "chk_shear": chk_shear, "chk_torsion": chk_torsion,
        "radio_critical": radio_critical, "radio_column": radio_column,
        "n1_hw": n1_hw, "d1_cb": d1_cb, "n2_hw": n2_hw, "d2_cb": d2_cb,
        "stirrup_plus": stirrup_plus,
        "btn_calc": btn_calc, "btn_note": btn_note,
        "grid_rect": grid_rect, "grid_hwnd": grid_hwnd,
        "results_log_hwnd": find_results_log_edit(rows, grid_rect),
    }


def click_control(hwnd):
    """同步点击（勿用于会弹出 rcba 模态框的操作，会死锁）"""
    if not is_valid_hwnd(hwnd):
        raise RuntimeError("控件句柄无效")
    win32gui.SendMessage(hwnd, win32con.BM_CLICK, 0, 0)
    time.sleep(SLEEP_SHORT)


def post_click_control(hwnd, label=""):
    """
    异步点击按钮。Calculate 等会弹出 rcba 模态框时必须用 PostMessage，
    否则 SendMessage 会一直阻塞到用户手动关弹窗，Python 无法继续点「是」。
    """
    if not is_valid_hwnd(hwnd):
        raise RuntimeError("控件句柄无效")
    if label:
        print(f"  → 点击: {label}")
    win32gui.PostMessage(hwnd, win32con.BM_CLICK, 0, 0)
    time.sleep(SLEEP_SHORT)


def set_edit_text(hwnd, text, main_hw=None, dismiss_rcba_label=None):
    """
    WM_SETTEXT 必须 SendMessage（PostMessage 会报错 1159）。
    若填写会触发 rcba（如 b/h 超限），在子线程 SendMessage，主线程并行点「是」避免死锁。
    """
    if not is_valid_hwnd(hwnd):
        raise RuntimeError("Edit 句柄无效")
    text = str(text)

    if main_hw and dismiss_rcba_label:
        done = threading.Event()
        errors = []

        def _send():
            try:
                win32gui.SendMessage(hwnd, win32con.WM_SETTEXT, 0, text)
            except Exception as exc:
                errors.append(exc)
            finally:
                done.set()

        threading.Thread(target=_send, daemon=True).start()
        ensure_rcba_dismissed(main_hw, timeout=8.0, label=dismiss_rcba_label)
        if not done.wait(timeout=10.0):
            raise RuntimeError(f"填写编辑框超时: {text}")
        if errors:
            raise errors[0]
    else:
        win32gui.SendMessage(hwnd, win32con.WM_SETTEXT, 0, text)

    time.sleep(SLEEP_SHORT)


def set_dimension_edit(main_hw, hwnd, value, label):
    """填写 b/h 等可能触发 rcba 几何警告的尺寸框，与 h 处理一致。"""
    set_edit_text(
        hwnd, value,
        main_hw=main_hw, dismiss_rcba_label=label,
    )
    if not _all_rcba_closed(main_hw):
        raise RuntimeError(f"rcba 警告框未关闭（{label}）")


def grid_lparam(client_x, client_y):
    return win32api.MAKELONG(int(client_x) & 0xFFFF, int(client_y) & 0xFFFF)


def focus_grid(main_hw, grid_hwnd):
    """把键盘焦点强制切到荷载表 GXWND"""
    if not is_valid_hwnd(grid_hwnd):
        return False
    try:
        win32gui.SetForegroundWindow(main_hw)
        win32gui.SetFocus(grid_hwnd)
    except Exception:
        pass
    time.sleep(0.12)
    focus = win32gui.GetFocus()
    for _ in range(8):
        if focus == grid_hwnd:
            return True
        if not focus:
            break
        focus = win32gui.GetParent(focus)
    return False


def click_grid_at(main_hw, grid_hwnd, screen_x, screen_y):
    """向 GXWND 发鼠标消息点击（比屏幕 mouse_event 更能拿到网格焦点）"""
    if not is_valid_hwnd(grid_hwnd):
        safe_click(main_hw, screen_x, screen_y)
        return
    cx, cy = win32gui.ScreenToClient(grid_hwnd, (int(screen_x), int(screen_y)))
    lp = grid_lparam(cx, cy)
    win32gui.SendMessage(grid_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lp)
    time.sleep(0.06)
    win32gui.SendMessage(grid_hwnd, win32con.WM_LBUTTONUP, 0, lp)
    time.sleep(0.12)
    focus_grid(main_hw, grid_hwnd)


def grid_send_key(main_hw, grid_hwnd, vk):
    """键盘消息直接发给 GXWND，避免落到 Results 文本框"""
    focus_grid(main_hw, grid_hwnd)
    if is_valid_hwnd(grid_hwnd):
        win32gui.SendMessage(grid_hwnd, win32con.WM_KEYDOWN, vk, 0)
        time.sleep(0.03)
        win32gui.SendMessage(grid_hwnd, win32con.WM_KEYUP, vk, 0)
    else:
        send_vk(vk)
    time.sleep(0.04)


def grid_send_text(main_hw, grid_hwnd, text):
    focus_grid(main_hw, grid_hwnd)
    for ch in str(text):
        if is_valid_hwnd(grid_hwnd):
            win32gui.SendMessage(grid_hwnd, win32con.WM_CHAR, ord(ch), 0)
        else:
            send_unicode_char(ch)
        time.sleep(0.025)


def safe_click(main_hw, x, y):
    """仅在 Expert 窗口范围内点击，不刷新/抢焦点（避免误关 Design 页）"""
    if not is_valid_hwnd(main_hw):
        raise RuntimeError("Expert 窗口句柄已失效，请重新打开 Design 页")
    ml, mt, mr, mb = win32gui.GetWindowRect(main_hw)
    if not (ml <= x <= mr and mt <= y <= mb):
        raise RuntimeError(
            f"坐标 ({int(x)},{int(y)}) 超出 Expert 窗口 ({ml},{mt},{mr},{mb})，已中止"
        )
    win32api.SetCursorPos((int(x), int(y)))
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    time.sleep(0.15)


def get_load_scroll_params(grid_rect=None) -> dict:
    """
    抗剪 Loads 表翻页参数（4 行可见，逻辑同 expert_try_biaxial 9 行翻页）。
    翻页：点最底可见行 Load type → Down；填表：用倒数第二行 Y（避免最底行落到滚动条）。
    """
    visible = SHEAR_LOAD_VISIBLE_ROWS
    return {
        "visible_rows": visible,
        "before_expand": visible,
        "scroll_click_row": SHEAR_LOAD_SCROLL_CLICK_ROW,
        "overflow_fill_row": SHEAR_LOAD_OVERFLOW_FILL_ROW,
    }


def effective_screen_row(logical_row_index: int, scroll_params: dict = None) -> int:
    """逻辑行超出可见数时，填表 Y 固定为最底可见行（翻页后新行所在位置）"""
    sp = scroll_params or get_load_scroll_params()
    if logical_row_index >= sp["before_expand"]:
        return sp["overflow_fill_row"]
    return logical_row_index


def load_col_screen_x(grid_rect, col_index):
    """0=Type, 1=V, 2=N"""
    gl, _, _, _ = grid_rect
    col_x = (LOAD_TYPE_COL_X, LOAD_V_COL_X, LOAD_N_COL_X)[col_index]
    return gl + col_x


def load_row_screen_y(grid_rect, screen_row_index):
    gl, gt, _, gb = grid_rect
    y = gt + LOAD_FIRST_ROW_Y + screen_row_index * LOAD_ROW_HEIGHT
    if y > gb - 8:
        return None
    return y


def load_cell_coords_at_screen_row(grid_rect, screen_row_index, col_index):
    """按指定屏幕行取单元格坐标（翻页点击与填表行可不同）"""
    y = load_row_screen_y(grid_rect, screen_row_index)
    if y is None:
        return None
    return load_col_screen_x(grid_rect, col_index), y


def load_cell_coords(grid_rect, logical_row_index, col_index, scroll_params=None):
    """三列：0=Type, 1=V, 2=N"""
    if scroll_params is None:
        scroll_params = get_load_scroll_params(grid_rect)
    screen_row = effective_screen_row(logical_row_index, scroll_params)
    return load_cell_coords_at_screen_row(grid_rect, screen_row, col_index)


def scroll_load_grid_down(main_hw, grid_hwnd, grid_rect, scroll_params=None):
    """
    第 5 组荷载起，每新增一行前先翻页一次（与 expert_try_biaxial 相同流程）：

    1. 物理点最底可见行（第 4 行）Load type
    2. focus_grid + grid_send_key Down
    3. 再点倒数第二行（第 3 行）Load type 作为填表锚点
    """
    if scroll_params is None:
        scroll_params = get_load_scroll_params(grid_rect)
    scroll_row = scroll_params["scroll_click_row"]
    fill_row = scroll_params["overflow_fill_row"]
    scroll_coords = load_cell_coords_at_screen_row(grid_rect, scroll_row, 0)
    fill_coords = load_cell_coords_at_screen_row(grid_rect, fill_row, 0)
    if not scroll_coords or not fill_coords:
        print("  ✗ 无法定位翻页/填表锚点（Load type）")
        return False
    main_hw = ensure_expert_visible(main_hw)
    print(
        f"  → 翻页：物理点第 {scroll_row + 1} 行 Load type "
        f"({int(scroll_coords[0])}, {int(scroll_coords[1])}) + Down"
    )
    safe_click(main_hw, scroll_coords[0], scroll_coords[1])
    time.sleep(SLEEP_CELL_CLICK)
    focus_grid(main_hw, grid_hwnd)
    grid_send_key(main_hw, grid_hwnd, win32con.VK_DOWN)
    time.sleep(SLEEP_CELL_SCROLL)
    print(
        f"  → 翻页后填表锚点：物理点第 {fill_row + 1} 行 Load type "
        f"({int(fill_coords[0])}, {int(fill_coords[1])})"
    )
    safe_click(main_hw, fill_coords[0], fill_coords[1])
    time.sleep(SLEEP_CELL_CLICK)
    focus_grid(main_hw, grid_hwnd)
    return True


def click_load_expand_button(main_hw, grid_rect):
    """备用：点荷载表右侧展开按钮（Down 无效时尝试）"""
    gl, gt, _, _ = grid_rect
    x, y = gl + LOAD_EXPAND_BTN_X, gt + LOAD_EXPAND_BTN_Y
    print(f"  → 点击荷载表展开按钮 ({int(x)}, {int(y)})")
    safe_click(main_hw, x, y)
    time.sleep(SLEEP_CELL_CLICK)


def scroll_load_grid_to_top(main_hw, grid_hwnd, grid_rect):
    """新工况填表前滚回顶部（物理点第 1 行 + SendInput Up）"""
    coords = load_cell_coords_at_screen_row(grid_rect, 0, 0)
    if not coords:
        return
    print("  → 滚回荷载表顶部")
    main_hw = ensure_expert_visible(main_hw)
    safe_click(main_hw, coords[0], coords[1])
    time.sleep(SLEEP_CELL_CLICK)
    for _ in range(15):
        send_vk(win32con.VK_UP)
        time.sleep(0.05)


def row_index_coords(grid_rect, row_index=0):
    """左侧行号列（点一下选中整行）"""
    gl, gt, _, gb = grid_rect
    y = gt + LOAD_FIRST_ROW_Y + row_index * LOAD_ROW_HEIGHT
    if y > gb - 8:
        return None
    return gl + LOAD_ROW_INDEX_COL_X, y


def format_cell_value(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    if isinstance(val, (int, float)):
        f = float(val)
        if f == int(f):
            return str(int(f))
        return f"{f:g}"
    return str(val).strip()


def format_load_number(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return "0.00"
    return f"{float(val):.2f}"


# ---------- 键盘输入（SendInput，比 keybd_event 更可靠） ----------
INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_UNICODE = 0x0004


class _KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", ctypes.c_ushort),
        ("wScan", ctypes.c_ushort),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]


class _INPUTunion(ctypes.Union):
    _fields_ = [("ki", _KEYBDINPUT)]


class _INPUT(ctypes.Structure):
    _anonymous_ = ("u",)
    _fields_ = [("type", ctypes.c_ulong), ("u", _INPUTunion)]


def _send_input(*events):
    arr = (_INPUT * len(events))(*events)
    ctypes.windll.user32.SendInput(len(events), arr, ctypes.sizeof(_INPUT))


def send_vk(vk):
    _send_input(
        _INPUT(type=INPUT_KEYBOARD, ki=_KEYBDINPUT(vk, 0, 0, 0, None)),
        _INPUT(type=INPUT_KEYBOARD, ki=_KEYBDINPUT(vk, 0, KEYEVENTF_KEYUP, 0, None)),
    )
    time.sleep(0.04)


def send_unicode_char(ch):
    code = ord(ch)
    _send_input(
        _INPUT(type=INPUT_KEYBOARD, ki=_KEYBDINPUT(0, code, KEYEVENTF_UNICODE, 0, None)),
        _INPUT(type=INPUT_KEYBOARD, ki=_KEYBDINPUT(0, code, KEYEVENTF_UNICODE | KEYEVENTF_KEYUP, 0, None)),
    )
    time.sleep(0.025)


def send_text_vk(text):
    """与 Delete 相同 SendInput 虚拟键路径，Expert 网格更认"""
    for ch in str(text):
        vk_scan = win32api.VkKeyScan(ch)
        if vk_scan == -1:
            send_unicode_char(ch)
            continue
        vk = vk_scan & 0xFF
        need_shift = (vk_scan >> 8) & 1
        if need_shift:
            _send_input(
                _INPUT(type=INPUT_KEYBOARD, ki=_KEYBDINPUT(win32con.VK_SHIFT, 0, 0, 0, None))
            )
        send_vk(vk)
        if need_shift:
            _send_input(
                _INPUT(type=INPUT_KEYBOARD, ki=_KEYBDINPUT(win32con.VK_SHIFT, 0, KEYEVENTF_KEYUP, 0, None))
            )


def send_text(text):
    send_text_vk(text)


def send_ctrl_a():
    _send_input(
        _INPUT(type=INPUT_KEYBOARD, ki=_KEYBDINPUT(win32con.VK_CONTROL, 0, 0, 0, None)),
        _INPUT(type=INPUT_KEYBOARD, ki=_KEYBDINPUT(ord("A"), 0, 0, 0, None)),
        _INPUT(type=INPUT_KEYBOARD, ki=_KEYBDINPUT(ord("A"), 0, KEYEVENTF_KEYUP, 0, None)),
        _INPUT(type=INPUT_KEYBOARD, ki=_KEYBDINPUT(win32con.VK_CONTROL, 0, KEYEVENTF_KEYUP, 0, None)),
    )
    time.sleep(0.06)


def send_ctrl_v():
    _send_input(
        _INPUT(type=INPUT_KEYBOARD, ki=_KEYBDINPUT(win32con.VK_CONTROL, 0, 0, 0, None)),
        _INPUT(type=INPUT_KEYBOARD, ki=_KEYBDINPUT(ord("V"), 0, 0, 0, None)),
        _INPUT(type=INPUT_KEYBOARD, ki=_KEYBDINPUT(ord("V"), 0, KEYEVENTF_KEYUP, 0, None)),
        _INPUT(type=INPUT_KEYBOARD, ki=_KEYBDINPUT(win32con.VK_CONTROL, 0, KEYEVENTF_KEYUP, 0, None)),
    )
    time.sleep(0.12)


def set_clipboard(text):
    import win32clipboard
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(str(text))
    win32clipboard.CloseClipboard()
    time.sleep(0.05)


class expert_keyboard_session:
    """整段荷载输入期间保持 Expert 键盘焦点（AttachThreadInput 不提前断开）"""

    def __init__(self, main_hw):
        self.main_hw = main_hw
        self.fg_tid = None
        self.target_tid = None
        self.attached = False

    def __enter__(self):
        self.main_hw = ensure_expert_visible(self.main_hw)
        fg = win32gui.GetForegroundWindow()
        self.target_tid = win32process.GetWindowThreadProcessId(self.main_hw)[0]
        if fg != self.main_hw:
            self.fg_tid = win32process.GetWindowThreadProcessId(fg)[0]
            if self.fg_tid and self.target_tid and self.fg_tid != self.target_tid:
                try:
                    win32process.AttachThreadInput(self.fg_tid, self.target_tid, True)
                    self.attached = True
                except Exception:
                    self.attached = False
            win32gui.SetForegroundWindow(self.main_hw)
            win32gui.BringWindowToTop(self.main_hw)
        time.sleep(SLEEP_KEYBOARD_SESSION)
        return self.main_hw

    def __exit__(self, *_):
        if not self.attached:
            return
        try:
            win32process.AttachThreadInput(self.fg_tid, self.target_tid, False)
        except Exception:
            pass
        self.attached = False


def bring_expert_to_front(main_hw):
    with expert_keyboard_session(main_hw) as hw:
        return hw


def get_cell_editor(grid_hwnd=None, grid_rect=None):
    """仅返回荷载表区域内的单元格小编辑框（排除下方 Results 大文本框）"""

    def in_grid_cell(edit_hw):
        if not grid_rect or not is_valid_hwnd(edit_hw):
            return False
        gl, gt, gr, gb = grid_rect
        l, t, r, b = win32gui.GetWindowRect(edit_hw)
        cx, cy = (l + r) // 2, (t + b) // 2
        w, h = r - l, b - t
        if w > 180 or h > 40:
            return False
        return gl <= cx <= gr and gt - 2 <= cy <= gb + 2

    for _ in range(10):
        time.sleep(0.06)
        focus = win32gui.GetFocus()
        if is_valid_hwnd(focus) and win32gui.GetClassName(focus) in EDIT_CLASSES:
            if in_grid_cell(focus):
                return focus

    if not is_valid_hwnd(grid_hwnd):
        return None
    found = []

    def walk(hw, _):
        if win32gui.GetClassName(hw) in EDIT_CLASSES and win32gui.IsWindowVisible(hw):
            if in_grid_cell(hw):
                found.append(hw)
        return True

    win32gui.EnumChildWindows(grid_hwnd, walk, None)
    return found[0] if found else None


def commit_cell_edit(main_hw, grid_hwnd):
    grid_send_key(main_hw, grid_hwnd, win32con.VK_RETURN)
    time.sleep(SLEEP_COMMIT_CELL)


def commit_cell_in_place(main_hw, grid_rect, grid_hwnd, logical_row_index, scroll_params):
    """提交当前格但不 Enter 换行（避免末行 N 按 Enter 误生成空行并错位）"""
    screen_row = effective_screen_row(logical_row_index, scroll_params)
    coords = load_cell_coords_at_screen_row(grid_rect, screen_row, 0)
    if coords and is_valid_hwnd(grid_hwnd):
        click_grid_at(main_hw, grid_hwnd, coords[0], coords[1])
    else:
        grid_send_key(main_hw, grid_hwnd, win32con.VK_TAB)
    time.sleep(SLEEP_COMMIT_CELL)


def commit_cell_tab(main_hw, grid_hwnd):
    """提交当前格并移到同行下一列（翻页后填表用）"""
    grid_send_key(main_hw, grid_hwnd, win32con.VK_TAB)
    time.sleep(0.2)


def write_cell_editor(grid_hwnd, grid_rect, text):
    edit = get_cell_editor(grid_hwnd, grid_rect)
    if not edit:
        return False
    win32gui.SendMessage(edit, win32con.WM_SETTEXT, 0, str(text))
    time.sleep(0.08)
    return win32gui.GetWindowText(edit).strip() == str(text).strip()


def type_in_grid_cell(main_hw, grid_hwnd, grid_rect, text):
    grid_send_key(main_hw, grid_hwnd, win32con.VK_F2)
    time.sleep(SLEEP_TYPE_CELL)
    if write_cell_editor(grid_hwnd, grid_rect, text):
        return
    grid_send_key(main_hw, grid_hwnd, win32con.VK_DELETE)
    time.sleep(0.06)
    grid_send_text(main_hw, grid_hwnd, text)


def _type_into_selected_cell(main_hw, grid_hwnd, grid_rect, col_index, text):
    if not str(text):
        return
    if col_index in (1, 2, 3):
        type_in_grid_cell(main_hw, grid_hwnd, grid_rect, text)
        print(f"    F2 + 网格内写入")
    else:
        grid_send_text(main_hw, grid_hwnd, text)


def click_and_type_cell(
    main_hw, grid_rect, grid_hwnd, logical_row_index, col_index, text, label,
    physical=False, scroll_params=None, commit_in_place=False,
):
    """点击单元格后输入；physical=True 时用物理鼠标（翻页后填表）"""
    if scroll_params is None:
        scroll_params = get_load_scroll_params(grid_rect)
    coords = load_cell_coords(grid_rect, logical_row_index, col_index, scroll_params)
    if not coords:
        print(f"  ✗ 第 {logical_row_index + 1} 行 {label} 坐标无效")
        return False
    screen_row = effective_screen_row(logical_row_index, scroll_params)
    click_mode = "物理点击" if physical else "点击"
    print(
        f"  → {click_mode} {label} ({int(coords[0])}, {int(coords[1])})  "
        f"输入: {text}  [逻辑行 {logical_row_index + 1} / 屏幕行 {screen_row + 1}]"
    )
    if physical:
        safe_click(main_hw, coords[0], coords[1])
        time.sleep(SLEEP_CELL_CLICK)
    else:
        click_grid_at(main_hw, grid_hwnd, coords[0], coords[1])
        time.sleep(SLEEP_CELL_CLICK)
    focus_grid(main_hw, grid_hwnd)

    _type_into_selected_cell(main_hw, grid_hwnd, grid_rect, col_index, text)
    if commit_in_place:
        commit_cell_in_place(main_hw, grid_rect, grid_hwnd, logical_row_index, scroll_params)
    else:
        commit_cell_edit(main_hw, grid_hwnd)
    return True


def _use_physical_for_row(logical_row: int, scroll_params: dict) -> bool:
    """最底可见行及翻页后填表行用物理点击"""
    return (
        logical_row >= scroll_params["before_expand"]
        or logical_row == scroll_params["before_expand"] - 1
    )


def ensure_expert_load_row(
    main_hw, grid_rect, grid_hwnd, logical_row, scroll_params,
):
    """逻辑行超出可见区时先翻页"""
    if logical_row >= scroll_params["before_expand"]:
        if not scroll_load_grid_down(main_hw, grid_hwnd, grid_rect, scroll_params):
            click_load_expand_button(main_hw, grid_rect)
        return True
    return False


def fill_all_load_rows(
    main_hw, load_rows, grid_rect, grid_hwnd, results_log_hw=None,
):
    blocked = None
    if is_valid_hwnd(results_log_hw):
        blocked = results_log_hw
        win32gui.EnableWindow(blocked, False)
        time.sleep(0.1)
    try:
        with expert_keyboard_session(main_hw) as main_hw:
            if not is_valid_hwnd(grid_hwnd):
                print("  警告: 未找到 GXWND 句柄，请确认在 Design 页")
            clear_loads_quick(main_hw, grid_rect, grid_hwnd)
            scroll_load_grid_to_top(main_hw, grid_hwnd, grid_rect)
            print("  → 清空完成，开始填入荷载...")
            time.sleep(SLEEP_LOAD_START)
            scroll_params = get_load_scroll_params(grid_rect)
            total = len(load_rows)
            print(
                f"  → Excel {total} 行 → Expert {total} 组荷载（每行 N+Qz 同行填入）"
                f"；界面可见 {scroll_params['visible_rows']} 行，"
                f"第 {scroll_params['before_expand'] + 1} 行起翻页"
                f"（翻页点第 {scroll_params['scroll_click_row'] + 1} 行，"
                f"填表 Y 第 {scroll_params['overflow_fill_row'] + 1} 行）"
            )
            for i, lr in enumerate(load_rows):
                print(
                    f"  荷载 {i + 1}/{total}: {lr['load_type']}  "
                    f"V={format_load_number(lr['V_kN'])}  "
                    f"N={format_load_number(lr['N_kN'])}"
                )
                fill_load_row(
                    main_hw, grid_rect, grid_hwnd,
                    lr["load_type"], lr["V_kN"], lr["N_kN"], i,
                    total, scroll_params,
                )
    finally:
        if blocked:
            win32gui.EnableWindow(blocked, True)
    verify_design_page(main_hw)
    print(f"  → 荷载表填写完成（共 {len(load_rows)} 组）")


def fill_load_row_overflow(
    main_hw, grid_rect, grid_hwnd, lt, v_str, n_str, row_index, scroll_params,
):
    if not scroll_load_grid_down(main_hw, grid_hwnd, grid_rect, scroll_params):
        raise RuntimeError(f"第 {row_index + 1} 行荷载翻页失败，已中止")
    screen_row = effective_screen_row(row_index, scroll_params)
    print(
        f"  --- 第 {row_index + 1} 行荷载: {lt}  V={v_str}  N={n_str}  "
        f"(屏幕行 {screen_row + 1}，物理点击逐格填表) ---"
    )
    for col, text, label, in_place in (
        (0, lt, "Load type", True),
        (1, v_str, "V", True),
        (2, n_str, "N", True),
    ):
        click_and_type_cell(
            main_hw, grid_rect, grid_hwnd, row_index, col, text, label,
            physical=True, scroll_params=scroll_params, commit_in_place=in_place,
        )


def fill_load_row(
    main_hw, grid_rect, grid_hwnd, load_type, v_val, n_val, row_index,
    total_rows, scroll_params,
):
    lt = format_cell_value(load_type)
    v_str = format_load_number(v_val)
    n_str = format_load_number(n_val)
    is_last_row = row_index >= total_rows - 1
    is_bottom_visible = row_index == scroll_params["before_expand"] - 1

    if row_index >= scroll_params["before_expand"]:
        fill_load_row_overflow(
            main_hw, grid_rect, grid_hwnd, lt, v_str, n_str, row_index, scroll_params,
        )
        print(f"  → 第 {row_index + 1} 行完成")
        return

    screen_row = effective_screen_row(row_index, scroll_params)
    use_physical = is_bottom_visible
    print(
        f"  --- 第 {row_index + 1} 行荷载: {lt}  V={v_str}  N={n_str}  "
        f"(屏幕行 {screen_row + 1}{'，物理点击' if use_physical else ''}) ---"
    )
    for col, text, label in (
        (0, lt, "Load type"),
        (1, v_str, "V"),
        (2, n_str, "N"),
    ):
        in_place = is_last_row or is_bottom_visible or col == 2
        click_and_type_cell(
            main_hw, grid_rect, grid_hwnd, row_index, col, text, label,
            physical=use_physical,
            scroll_params=scroll_params,
            commit_in_place=in_place,
        )
    print(f"  → 第 {row_index + 1} 行完成")


def verify_design_page(main_hw):
    get_design_page_hwnd(main_hw)


def ensure_checkbox_unchecked(hwnd):
    if not is_valid_hwnd(hwnd):
        return
    if win32gui.SendMessage(hwnd, win32con.BM_GETCHECK, 0, 0) == win32con.BST_CHECKED:
        win32gui.PostMessage(hwnd, win32con.BM_CLICK, 0, 0)
    time.sleep(SLEEP_SHORT)


def ensure_radio_checked(hwnd):
    if not is_valid_hwnd(hwnd):
        return
    if win32gui.SendMessage(hwnd, win32con.BM_GETCHECK, 0, 0) != win32con.BST_CHECKED:
        win32gui.PostMessage(hwnd, win32con.BM_CLICK, 0, 0)
    time.sleep(SLEEP_SHORT)


def set_combobox_value(hwnd, value, main_hw=None):
    if not is_valid_hwnd(hwnd):
        return False
    text = str(int(float(value))) if value is not None else ""
    if not text:
        return False

    idx = -1
    for i, item in enumerate(_combo_list_items(hwnd, open_drop=True)):
        bare = re.sub(r"^[φΦ]\s*", "", item).strip()
        if item == text or bare == text:
            idx = i
            break
    if idx < 0:
        idx = win32gui.SendMessage(hwnd, win32con.CB_FINDSTRINGEXACT, -1, text)
    if idx < 0:
        idx = win32gui.SendMessage(hwnd, win32con.CB_FINDSTRING, -1, text)

    if idx >= 0:
        win32gui.SendMessage(hwnd, win32con.CB_SETCURSEL, idx, 0)
        _notify_combo_selchange(hwnd)
        time.sleep(SLEEP_SHORT)
        return True

    if main_hw:
        try:
            rect = win32gui.GetWindowRect(hwnd)
            safe_click(main_hw, (rect[0] + rect[2]) // 2, (rect[1] + rect[3]) // 2)
            time.sleep(0.12)
            send_ctrl_a()
            send_text(text)
            send_vk(win32con.VK_RETURN)
            _notify_combo_selchange(hwnd)
            time.sleep(SLEEP_SHORT)
            return True
        except Exception:
            pass
    return False


def _is_control_enabled(hwnd) -> bool:
    return bool(hwnd and is_valid_hwnd(hwnd) and win32gui.IsWindowEnabled(hwnd))


def wait_control_enabled(hwnd, timeout: float = 3.0) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        if _is_control_enabled(hwnd):
            return True
        time.sleep(0.05)
    return False


def _bar_cross_section_area_mm2(diameter: int) -> float:
    """单根钢筋面积 (mm²)，d 为整数直径 (mm)。"""
    d = int(diameter)
    return math.pi * d * d / 4.0


def convert_stirrup_second_group_to_first(
    n1, d_bar1, n2, d_bar2,
) -> tuple:
    """
    将 n2×φ(d_bar2) 折算为若干根 φ(d_bar1) 并入 n1。
    向下取整保证折算面积 ≤ n2×d_bar2 原面积（偏安全，不超配）。
  """
    n1_i = int(_safe_int(n1, 0))
    d1_i = int(_parse_bar_diameter(d_bar1, 6))
    n2_i = int(_safe_int(n2, 0))
    if n2_i <= 0:
        return n1_i, d1_i, {"converted": False}

    d2_i = int(_parse_bar_diameter(d_bar2, 6))
    area_second = n2_i * _bar_cross_section_area_mm2(d2_i)
    area_per_d1 = _bar_cross_section_area_mm2(d1_i)
    extra_bars = int(math.floor(area_second / area_per_d1))
    converted_area = extra_bars * area_per_d1

    info = {
        "converted": True,
        "original": f"{n1_i}φ{d1_i}+{n2_i}φ{d2_i}",
        "equivalent": f"{n1_i + extra_bars}φ{d1_i}",
        "extra_bars": extra_bars,
        "area_second_mm2": area_second,
        "area_converted_mm2": converted_area,
    }
    return n1_i + extra_bars, d1_i, info


def fill_stirrup_reinforcement(ctrls, n1, d_bar1, n2, d_bar2, main_hw=None):
    """
    Results 区 Cadres：仅填 n1 + φ1。
    有第二组时按面积折算并入 n1（折算面积 ≤ 原第二组面积）。
    """
    n1_i, d1_i = _safe_int(n1, 0), _parse_bar_diameter(d_bar1, 6)
    n2_i, d2_i = _safe_int(n2, 0), _parse_bar_diameter(d_bar2, 6)
    desc = f"{n1_i}φ{d1_i}" + (f"+{n2_i}φ{d2_i}" if n2_i else "")
    print(f"  → 箍筋 Cadres: {desc}")

    if not main_hw:
        print("  ✗ 缺少主窗口句柄，无法填写箍筋")
        return

    fill_n1, fill_d1, conv = convert_stirrup_second_group_to_first(
        n1_i, d1_i, n2_i, d2_i,
    )
    if conv.get("converted"):
        print(
            f"  → 第二组面积折算: {conv['original']} → {conv['equivalent']} "
            f"(+{conv['extra_bars']}根φ{fill_d1}, "
            f"折算{conv['area_converted_mm2']:.1f} mm² "
            f"≤ 原{conv['area_second_mm2']:.1f} mm²)"
        )

    fill_stirrup_sequential(main_hw, ctrls, fill_n1, fill_d1)


def apply_shear_options(ctrls):
    if ctrls.get("chk_shear"):
        ensure_checkbox_checked(ctrls["chk_shear"])
    if ctrls.get("chk_torsion"):
        ensure_checkbox_unchecked(ctrls["chk_torsion"])
    if ctrls.get("chk_seismic"):
        ensure_checkbox_checked(ctrls["chk_seismic"])
    if ctrls.get("radio_critical"):
        ensure_radio_checked(ctrls["radio_critical"])
    if ctrls.get("radio_column"):
        ensure_radio_checked(ctrls["radio_column"])


def ensure_checkbox_checked(hwnd):
    if not is_valid_hwnd(hwnd):
        return
    if win32gui.SendMessage(hwnd, win32con.BM_GETCHECK, 0, 0) != win32con.BST_CHECKED:
        win32gui.PostMessage(hwnd, win32con.BM_CLICK, 0, 0)
    time.sleep(SLEEP_SHORT)


def _is_rcba_yes_button(text):
    """rcba 对话框上应点的「继续」按钮（几何/弯矩/轴力等超限警告）"""
    t = (text or "").strip()
    if not t:
        return False
    if t in ("确定", "OK", "Yes", "是", "是(Y)", "是 (Y)", "是（Y）", "是(&Y)"):
        return True
    return t.startswith("是") and "否" not in t


# rcba Static 文案片段（几何 b/h、弯矩 My/Mz、轴力 N 等超限警告共用类似句式）
_RCBA_HINT_SNIPPETS = (
    # 通用句式（各类超限弹窗常见）
    "do you wish to continue",
    "wish to continue",
    "continue calculations",
    "continue calculation",
    "condition is not fulfilled",
    "not fulfilled",
    "may be incorrect",
    "incorrect",
    "currently set values",
    "currently set value",
    "following condition",
    "warning",
    "exceed",
    "exceeded",
    "maximum",
    "allowable",
    "limit",
    "capacity",
    "less than",
    "greater than",
    "out of",
    "not valid",
    "invalid",
    # 几何尺寸 b / h
    "geometrical data",
    "geometrical",
    "geometry",
    "200 cm",
    "<= 200",
    "b <=",
    "h <=",
    "width",
    "breadth",
    "height",
    " cm",
    "cross-section",
    "cross section",
    "section dimension",
    # 荷载 / 弯矩 / 轴力
    "loading data",
    "load data",
    "loading",
    "loads",
    "moment",
    "moments",
    "bending moment",
    "bending",
    "biaxial",
    "my",
    "mz",
    "my =",
    "mz =",
    "n-m",
    "n m",
    "nm interaction",
    "interaction diagram",
    "interaction",
    "axial load",
    "axial force",
    "axial",
    "compression",
    "tension",
    "shear",
    "torque",
    "kn·m",
    "knm",
    "kn*m",
    "kn m",
    "kN·m",
    "kN*m",
    # 中文
    "继续",
    "是否继续",
    "继续计算",
    "几何",
    "不正确",
    "不符合",
    "条件",
    "当前值",
    "目前",
    "宽度",
    "高度",
    "截面",
    "尺寸",
    "弯矩",
    "轴力",
    "荷载",
    "内力",
    "超出",
    "超过",
    "过大",
    "极限",
    "承载",
    "限制",
    "警告",
    "双轴",
    "弯曲",
    "轴向",
    "剪力",
    "扭矩",
)


def _static_text_is_rcba_hint(text: str) -> bool:
    t = (text or "").lower()
    if not t.strip():
        return False
    if any(snippet in t for snippet in _RCBA_HINT_SNIPPETS):
        return True
    # Expert 弯矩/轴力警告里常见纯数值行（如 My = 7632.44 kN·m）
    if re.search(r"\b(my|mz|n)\s*[=:]", t):
        return True
    if re.search(r"\d+\.?\d*\s*(kn|knm|kn·m|kn\*m)", t):
        return True
    return False


def _is_rcba_warning_dialog(dlg):
    """带「是/否」的 rcba 警告框（b/h/弯矩/轴力等超限），避免误认其他 #32770"""
    if not is_valid_hwnd(dlg):
        return False
    if win32gui.GetClassName(dlg) != "#32770":
        return False
    if win32gui.GetWindowText(dlg) != "rcba":
        return False
    if not is_valid_hwnd(win32gui.GetDlgItem(dlg, win32con.IDYES)):
        return False
    if not is_valid_hwnd(win32gui.GetDlgItem(dlg, win32con.IDNO)):
        return False
    has_hint = False

    def on_static(hw, _):
        nonlocal has_hint
        if win32gui.GetClassName(hw) != "Static":
            return True
        if _static_text_is_rcba_hint(win32gui.GetWindowText(hw)):
            has_hint = True
        return True

    win32gui.EnumChildWindows(dlg, on_static, None)
    return has_hint


def _find_rcba_dialogs(main_hw=None):
    """仅枚举顶层可见的 rcba 几何警告框（不遍历 Expert 子树，避免误匹配）"""
    found = []
    win32gui.EnumWindows(
        lambda hw, arr: arr.append(hw) or True
        if _is_rcba_warning_dialog(hw) and win32gui.IsWindowVisible(hw)
        else True,
        found,
    )
    return found


def _all_rcba_closed(main_hw=None):
    return len(_find_rcba_dialogs(main_hw)) == 0


def _set_foreground_window(hwnd):
    """跨线程 SetForegroundWindow（Python 脚本线程 → UI 线程）"""
    if not is_valid_hwnd(hwnd):
        return
    try:
        if win32gui.GetForegroundWindow() == hwnd:
            return
        fg_tid = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())[0]
        target_tid = win32process.GetWindowThreadProcessId(hwnd)[0]
        attached = False
        if fg_tid and target_tid and fg_tid != target_tid:
            try:
                win32process.AttachThreadInput(fg_tid, target_tid, True)
                attached = True
            except Exception:
                attached = False
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            win32gui.BringWindowToTop(hwnd)
            win32gui.SetForegroundWindow(hwnd)
        finally:
            if attached:
                try:
                    win32process.AttachThreadInput(fg_tid, target_tid, False)
                except Exception:
                    pass
    except Exception:
        pass


def _focus_rcba_dialog(dlg):
    """
    只置前 rcba 模态框本身。
    切勿在弹框已显示时 SetForegroundWindow(Expert 主窗口)，会把模态框顶到后面导致点不中。
    """
    _set_foreground_window(dlg)
    time.sleep(0.1)


def _rcba_yes_button(dlg):
    yes_btn = win32gui.GetDlgItem(dlg, win32con.IDYES)
    if is_valid_hwnd(yes_btn):
        return yes_btn
    for hw in _enum_child_buttons(dlg):
        cid = win32gui.GetWindowLong(hw, win32con.GWL_ID)
        if cid == win32con.IDYES or _is_rcba_yes_button(win32gui.GetWindowText(hw)):
            return hw
    return None


def _physical_click(hwnd):
    l, t, r, b = win32gui.GetWindowRect(hwnd)
    cx, cy = (l + r) // 2, (t + b) // 2
    win32api.SetCursorPos((cx, cy))
    time.sleep(0.06)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.06)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)


def _safe_physical_click(hwnd):
    """物理点击；对话框关闭后子控件句柄会失效，须容错。"""
    if not is_valid_hwnd(hwnd):
        return False
    try:
        _physical_click(hwnd)
        return True
    except Exception:
        return False


def _dialog_still_open(dlg):
    return is_valid_hwnd(dlg) and _is_rcba_warning_dialog(dlg)


def _click_rcba_yes(dlg, main_hw=None):
    """
    关闭单个 rcba 警告框（b/h 超限等）。
    WM_COMMAND 成功后对话框与子按钮句柄会立即销毁，不可复用 yes_btn。
    """
    if not is_valid_hwnd(dlg):
        return True
    if not _is_rcba_warning_dialog(dlg):
        return False

    yes_btn = _rcba_yes_button(dlg)
    if not is_valid_hwnd(yes_btn):
        print("  ✗ rcba：未找到「是」按钮")
        return False

    _focus_rcba_dialog(dlg)
    win32gui.SendMessage(dlg, win32con.WM_COMMAND, win32con.IDYES, yes_btn)
    time.sleep(0.3)
    if not _dialog_still_open(dlg):
        return True

    yes_btn = _rcba_yes_button(dlg)
    if is_valid_hwnd(yes_btn):
        _focus_rcba_dialog(dlg)
        _safe_physical_click(yes_btn)
        time.sleep(0.25)
    if not _dialog_still_open(dlg):
        return True

    yes_btn = _rcba_yes_button(dlg)
    if is_valid_hwnd(yes_btn):
        win32gui.SendMessage(yes_btn, win32con.BM_CLICK, 0, 0)
        time.sleep(0.25)

    return not _dialog_still_open(dlg)


def ensure_rcba_dismissed(main_hw=None, timeout=5.0, label=""):
    """循环关闭全部 rcba（b、h 各超限时会连续弹出多个警告框）"""
    deadline = time.time() + timeout
    closed = 0
    while time.time() < deadline:
        dialogs = _find_rcba_dialogs(main_hw)
        if not dialogs:
            if closed:
                print(f"  → 共关闭 {closed} 个 rcba 警告框，继续后续步骤")
            else:
                print("  → rcba 警告框已关闭，继续后续步骤")
            time.sleep(SLEEP_SHORT)
            return True

        for dlg in dialogs:
            if not is_valid_hwnd(dlg):
                continue
            print(
                f"  → 发现 rcba 几何警告框（{label or '自动处理'}）"
                f" hwnd={dlg}，点击「是」..."
            )
            if _click_rcba_yes(dlg, main_hw):
                closed += 1
            time.sleep(0.12)

        if _all_rcba_closed(main_hw):
            if closed:
                print(f"  → 共关闭 {closed} 个 rcba 警告框，继续后续步骤")
            else:
                print("  → rcba 警告框已关闭，继续后续步骤")
            time.sleep(SLEEP_SHORT)
            return True

        print("  → 仍有 rcba 警告框，继续处理...")
        time.sleep(SLEEP_RCBA_POLL)

    remaining = _find_rcba_dialogs(main_hw)
    if remaining:
        print(f"  ✗ rcba 在 {timeout}s 内未能全部关闭（剩余 {len(remaining)} 个）")
        return False
    return True


def dismiss_rcba_dialog(main_hw=None):
    """仅在已确认存在 rcba 警告框时处理（填表过程中勿调用）"""
    if not _find_rcba_dialogs(main_hw):
        return False
    return ensure_rcba_dismissed(main_hw, timeout=3.0, label="")


def _enum_child_buttons(parent_hw):
    buttons = []

    def walk(hw, _):
        if win32gui.GetClassName(hw) in BUTTON_CLASS:
            buttons.append(hw)
        return True

    win32gui.EnumChildWindows(parent_hw, walk, None)
    return buttons


def wait_after_calculate(main_hw):
    """Calculate 后：发现 rcba 即点「是」"""
    ensure_expert_visible(main_hw)
    return ensure_rcba_dismissed(main_hw, timeout=SLEEP_CALC_MAX, label="Calculate后")


def clear_loads_quick(main_hw, grid_rect, grid_hwnd, times=LOAD_CLEAR_TIMES):
    """点第 1 行左侧行号 → Delete 清空整行，重复 times 次"""
    print(f"  清空荷载：点行号 1 + Delete × {times}")
    coords = row_index_coords(grid_rect, 0)
    if not coords:
        return
    for _ in range(times):
        click_grid_at(main_hw, grid_hwnd, coords[0], coords[1])
        time.sleep(SLEEP_GRID_CLEAR)
        grid_send_key(main_hw, grid_hwnd, win32con.VK_DELETE)
        time.sleep(SLEEP_GRID_CLEAR)


def _is_rc_note_window_title(title: str) -> bool:
    """仅匹配 Expert Note 生成的 rc_note.rtf 窗口标题，避免误关其他 Word 文档。"""
    t = (title or "").strip().lower().replace("\\", "/")
    if not t:
        return False
    return "rc_note.rtf" in t or t == "rc_note" or t.startswith("rc_note ")


def close_rc_note_viewer():
    """仅关闭正在查看 rc_note.rtf 的窗口（Word/写字板），不关其他 Word 文件。"""

    def close_hw(hw, _):
        if not win32gui.IsWindowVisible(hw):
            return True
        title = win32gui.GetWindowText(hw)
        if TARGET_WINDOW in title:
            return True
        if not _is_rc_note_window_title(title):
            return True
        print(f"  → 关闭 rc_note 查看器: {title[:60]}")
        win32gui.PostMessage(hw, win32con.WM_CLOSE, 0, 0)
        return True

    win32gui.EnumWindows(close_hw, None)
    time.sleep(0.3)


def find_rc_note_file(timeout=8, since_mtime=0):
    """等待 Expert 写出 rc_note.rtf（优先 Autodesk Output 目录）"""
    candidates = [
        EXPERT_RTF_PATH,
        os.path.join(SCRIPT_DIR, "rc_note.rtf"),
        os.path.join(os.getcwd(), "rc_note.rtf"),
    ]
    end = time.time() + timeout
    while time.time() < end:
        for p in candidates:
            if os.path.isfile(p) and os.path.getsize(p) > 0:
                if since_mtime <= 0 or os.path.getmtime(p) >= since_mtime - 0.5:
                    return p
        time.sleep(0.3)
    return None


def _safe_filename_part(text: str) -> str:
    s = str(text).strip().replace(os.sep, "_")
    for ch in ':*?"<>|':
        s = s.replace(ch, "_")
    return s


def rtf_save_name(region: str, code: str = "", append_code: bool = False) -> str:
    """
    RTF 另存名：默认 {区域}_shear。
    同一区域存在多个不同编号工况时，追加编号：{区域}_{编号}_shear。
    """
    base = _safe_filename_part(region)
    if append_code and code:
        return f"{base}_{_safe_filename_part(code)}_shear"
    return f"{base}_shear"


def assign_rtf_names(cases: List[dict]) -> None:
    """按区域统计工况数，为每条 case 设置 rtf_name。"""
    region_counts: Dict[str, int] = {}
    for c in cases:
        region_counts[c["region"]] = region_counts.get(c["region"], 0) + 1
    for c in cases:
        c["rtf_name"] = rtf_save_name(
            c["region"], c["code"], region_counts[c["region"]] > 1,
        )


def save_rtf(note_id):
    """从 Expert 输出目录复制 rc_note.rtf，按编号另存到 OUTPUT_RTF_DIR"""
    os.makedirs(OUTPUT_RTF_DIR, exist_ok=True)
    safe_id = str(note_id).strip().replace(os.sep, "_")
    dest = os.path.join(OUTPUT_RTF_DIR, f"{safe_id}.rtf")

    before = time.time()
    close_rc_note_viewer()
    dismiss_rcba_dialog()

    src = find_rc_note_file(timeout=SLEEP_NOTE + 5, since_mtime=before)
    if not src:
        src = find_rc_note_file(timeout=3, since_mtime=0)

    if src:
        shutil.copy2(src, dest)
        print(f"  RTF 已保存: {dest}")
    else:
        print(f"  警告: 未找到 rc_note.rtf，请检查 {EXPERT_RTF_PATH}")

    close_rc_note_viewer()
    dismiss_rcba_dialog()
    return dest if os.path.isfile(dest) else ""


def _norm_header_key(name) -> str:
    s = str(name).strip().lower()
    s = s.replace("（", "(").replace("）", ")").replace("·m", "m").replace("·", "")
    s = re.sub(r"\s+", "", s)
    return s


def _map_shear_column(col_name: str) -> str:
    key = _norm_header_key(col_name)
    exact = {
        "区域": "region",
        "编号": "code",
        "名称": "load_name",
        "类型": "force_type",
        "计算项": "calc_item",
        "n(kn)": "N_kN",
        "n_kn": "N_kN",
        "qz(knm)": "Qz_kNm",
        "qz_knm": "Qz_kNm",
        "bcm": "b_cm",
        "hcm": "h_cm",
        "d1cm": "d1_cm",
        "d2cm": "d2_cm",
        "n1": "n1",
        "n2": "n2",
        "dbar1": "d_bar1",
        "dbar2": "d_bar2",
        "dbar": "d_bar1",
        "箍筋": "stirrup",
        "cadres": "stirrup",
        "配筋": "stirrup",
    }
    if key in exact:
        return exact[key]
    if key.startswith("d_bar") or key.startswith("dbar"):
        if key.endswith("2") or key.endswith("02"):
            return "d_bar2"
        return "d_bar1"
    return str(col_name).strip()


def normalize_shear_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename = {}
    for c in df.columns:
        mapped = _map_shear_column(c)
        if mapped == "d_bar1" and "d_bar1" in rename.values():
            mapped = "d_bar2"
        rename[c] = mapped
    return df.rename(columns=rename)


def forward_fill_merged_cells(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in out.columns:
        out[col] = out[col].ffill()
    return out


def _clean_load_type(val) -> str:
    """保留「ULS MAX」等完整名称，便于区分 MAX/MIN 包络。"""
    if pd.isna(val):
        return "ULS"
    s = str(val).strip()
    if not s:
        return "ULS"
    parts = s.split()
    head = parts[0].upper()
    if head in ("ULS", "ALS", "SLS") and len(parts) >= 2:
        return f"{head} {parts[1].upper()}"
    return head


def _safe_int(val, default=0) -> int:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return default
    try:
        return int(float(val))
    except (TypeError, ValueError):
        return default


STIRRUP_SINGLE_RE = re.compile(
    r"^(\d+)\s*[φΦ*×]\s*(\d+)$",
    re.IGNORECASE,
)
STIRRUP_COMBINED_RE = re.compile(
    r"^(\d+)\s*[φΦ*×]\s*(\d+)\s*\+\s*(\d+)\s*[φΦ*×]\s*(\d+)$",
    re.IGNORECASE,
)


def _looks_like_stirrup_spec(val) -> bool:
    """合并写法须含 φ 或 +，纯数字列值不走合并解析。"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return False
    s = str(val).strip()
    if not s:
        return False
    return any(ch in s for ch in ("φ", "Φ", "Ф", "ø", "Ø", "*", "×", "+"))


def _parse_bar_diameter(val, default=6) -> int:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return default
    s = str(val).strip()
    m = re.search(r"(\d+)", s)
    return int(m.group(1)) if m else _safe_int(val, default)


def parse_stirrup_reinforcement(text) -> Optional[dict]:
    """解析「2φ16+10φ12」→ n1/d_bar1/n2/d_bar2（不含 φ/+ 的纯数字返回 None）。"""
    if text is None or (isinstance(text, float) and pd.isna(text)):
        return None
    s = str(text).strip()
    if not s or not _looks_like_stirrup_spec(s):
        return None
    for ch, repl in (("Φ", "φ"), ("Ф", "φ"), ("ø", "φ"), ("Ø", "φ"), ("*", "φ"), ("×", "φ")):
        s = s.replace(ch, "φ")
    s = re.sub(r"\s+", "", s)

    m = STIRRUP_COMBINED_RE.match(s)
    if m:
        return {
            "n1": int(m.group(1)),
            "d_bar1": int(m.group(2)),
            "n2": int(m.group(3)),
            "d_bar2": int(m.group(4)),
        }
    m = STIRRUP_SINGLE_RE.match(s)
    if m:
        return {
            "n1": int(m.group(1)),
            "d_bar1": int(m.group(2)),
            "n2": 0,
            "d_bar2": 6,
        }
    return None


def extract_stirrup_params(row) -> dict:
    """从 Excel 行读取箍筋：优先分列 n1/d_bar1/n2/d_bar2；合并列须含 φ 或 +。"""
    if "stirrup" in row.index and pd.notna(row.get("stirrup")):
        parsed = parse_stirrup_reinforcement(row["stirrup"])
        if parsed:
            return parsed

    n1_raw = row.get("n1")
    if _looks_like_stirrup_spec(n1_raw):
        parsed = parse_stirrup_reinforcement(n1_raw)
        if parsed:
            return parsed

    n2_raw = row.get("n2")
    d2_raw = row.get("d_bar2")
    n2_val = _safe_int(n2_raw, 0)
    d2_val = _parse_bar_diameter(d2_raw, 0) if pd.notna(d2_raw) else 0

    return {
        "n1": _safe_int(n1_raw, 2),
        "d_bar1": _parse_bar_diameter(row.get("d_bar1"), 6),
        "n2": n2_val,
        "d_bar2": _parse_bar_diameter(d2_raw, 6) if n2_val > 0 else 6,
    }


def extract_stirrup_from_group(group: pd.DataFrame) -> dict:
    """从分组中取第一条含 d_bar1 的行作为箍筋参数（避免 Qz 行覆盖 N 行）。"""
    for _, row in group.iterrows():
        d1 = row.get("d_bar1")
        if pd.notna(d1) and _parse_bar_diameter(d1, 0) > 0:
            return extract_stirrup_params(row)
    return extract_stirrup_params(group.iloc[0])


def build_loads_from_group(group: pd.DataFrame) -> List[dict]:
    """
    抗剪 sheet 每个点位通常 8 行 Excel，每行独立一组荷载。
    每行的 N(kN) 与 Qz(kNm) 均取自同一行，填入 Expert 一行（共 8 组）。
    """
    name_col = "load_name" if "load_name" in group.columns else "名称"
    loads: List[dict] = []
    for _, r in group.iterrows():
        name_s = str(r.get(name_col, "")).strip()
        if not name_s or name_s.lower() == "nan":
            continue
        n_val = r.get("N_kN")
        v_val = r.get("Qz_kNm")
        if pd.isna(n_val) and pd.isna(v_val):
            continue
        loads.append({
            "load_type": _clean_load_type(name_s),
            "N_kN": float(n_val) if pd.notna(n_val) else 0.0,
            "V_kN": float(v_val) if pd.notna(v_val) else 0.0,
        })
    return loads


def read_shear_cases(path: str) -> List[dict]:
    """读取 input_VN.xlsx「抗剪」sheet，按区域+编号分组为计算工况。"""
    df = pd.read_excel(path, sheet_name=INPUT_SHEET, dtype=object)
    df = normalize_shear_columns(df)
    df = forward_fill_merged_cells(df)

    for col in ("region", "code", "b_cm", "h_cm", "d1_cm"):
        if col not in df.columns:
            raise RuntimeError(f"Excel 缺少列: {col}（sheet={INPUT_SHEET}）")

    group_cols = ["region", "code"]
    if "calc_item" in df.columns:
        group_cols = ["calc_item", "region", "code"]

    cases: List[dict] = []
    for key, group in df.groupby(group_cols, sort=False):
        row0 = group.iloc[0]
        region = str(row0["region"]).strip()
        code = str(row0["code"]).strip()
        calc_item = str(row0.get("calc_item", "")).strip() if "calc_item" in group.columns else ""
        id_parts = [p for p in (calc_item, region, code) if p]
        case_id = "_".join(id_parts)

        loads = build_loads_from_group(group)
        if not loads:
            continue

        stirrup = extract_stirrup_from_group(group)
        cases.append({
            "ID": case_id,
            "region": region,
            "code": code,
            "b_cm": float(row0["b_cm"]),
            "h_cm": float(row0["h_cm"]),
            "d1_cm": float(row0["d1_cm"]),
            "n1": stirrup["n1"],
            "d_bar1": stirrup["d_bar1"],
            "n2": stirrup["n2"],
            "d_bar2": stirrup["d_bar2"],
            "loads": loads,
        })
    assign_rtf_names(cases)
    return cases


def run_one(case: dict, main_hw):
    main_hw = ensure_expert_visible(main_hw)
    ctrls = build_control_map(main_hw)
    grid_rect = ctrls["grid_rect"]
    grid_hwnd = ctrls.get("grid_hwnd")
    loads = case["loads"]

    print("  → 选项: shear / 非 torsion / Seismic / Critical / Column")
    apply_shear_options(ctrls)

    fill_all_load_rows(
        main_hw, loads, grid_rect, grid_hwnd, ctrls.get("results_log_hwnd"),
    )

    fill_stirrup_reinforcement(
        ctrls, case["n1"], case["d_bar1"], case["n2"], case["d_bar2"],
        main_hw=main_hw,
    )

    print("  → 填写截面尺寸 b, h, d（保护层 d1）...")
    set_dimension_edit(main_hw, ctrls["edit_b"], case["b_cm"], "填写 b 后")
    set_dimension_edit(main_hw, ctrls["edit_h"], case["h_cm"], "填写 h 后")
    set_edit_text(ctrls["edit_d"], case["d1_cm"])

    post_click_control(ctrls["btn_calc"], "Calculate")
    time.sleep(0.35)
    if not wait_after_calculate(main_hw):
        raise RuntimeError("rcba 警告框未关闭，已中止")

    expert_out = read_shear_expert_results(
        main_hw, grid_rect, ctrls.get("results_log_hwnd"),
    )

    if not _all_rcba_closed(main_hw):
        raise RuntimeError("rcba 警告框仍打开，无法继续 Note")
    close_rc_note_viewer()
    post_click_control(ctrls["btn_note"], "Note")
    time.sleep(SLEEP_NOTE)
    new_rtf = save_rtf(case["rtf_name"])

    return build_result_summary_row(case, expert_out, new_rtf)


def create_sample_input(path: str):
    rows = [
        {
            "区域": "C0底 1#", "编号": "6017 j端", "名称": "ULS MAX",
            "类型": "N", "N(kN)": 1161.14, "Qz(kNm)": -668.51,
            "b_cm": 100, "h_cm": 250, "d1_cm": 9, "d2_cm": 9,
            "n1": 2, "d_bar1": 6, "n2": 0, "d_bar2": 6,
        },
        {
            "区域": "C0底 1#", "编号": "6017 j端", "名称": "ULS MIN",
            "类型": "N", "N(kN)": 800.0, "Qz(kNm)": -200.0,
            "b_cm": 100, "h_cm": 250, "d1_cm": 9, "d2_cm": 9,
            "n1": 2, "d_bar1": 6, "n2": 0, "d_bar2": 6,
        },
        {
            "区域": "C0底 1#", "编号": "6017 j端", "名称": "ALS MAX",
            "类型": "Qz", "N(kN)": 900.0, "Qz(kNm)": -500.0,
            "b_cm": 100, "h_cm": 250, "d1_cm": 9, "d2_cm": 9,
            "n1": 2, "d_bar1": 6, "n2": 0, "d_bar2": 6,
        },
    ]
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame(rows).to_excel(writer, sheet_name=INPUT_SHEET, index=False)


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] in ("--calibrate", "--test-coords"):
        from calibrate_load_shear import run_calibrate, run_test
        if sys.argv[1] == "--calibrate":
            run_calibrate()
        else:
            run_test()
        raise SystemExit(0)

    print("=" * 50)
    print("Expert 抗剪验算 (Shear and torsion)")
    print("运行前请：打开 Expert → Design 页 → 窗口不要最小化")
    print("运行期间请勿操作鼠标键盘")
    print(f"输入: {INPUT_EXCEL}（sheet={INPUT_SHEET}）")
    print(f"RTF 另存目录: {OUTPUT_RTF_DIR}")
    print(f"结果汇总: {OUTPUT_EXCEL}")
    if os.path.isfile(LOAD_COORDS_FILE):
        print(f"荷载坐标: 已加载 {LOAD_COORDS_FILE}")
        print(
            f"  Type={LOAD_TYPE_COL_X}  V={LOAD_V_COL_X}  N={LOAD_N_COL_X}  "
            f"首行Y={LOAD_FIRST_ROW_Y}  行高={LOAD_ROW_HEIGHT}"
        )
        print("  箍筋: 点击+Enter→DlgTab→F2（PostMessage Tab，非全局 SendInput）")
    else:
        print("荷载坐标: 未校准，请运行  python calibrate_load_shear.py")
        print("          或  python expert_try_shear.py --calibrate")
    print("=" * 50)

    if not os.path.exists(INPUT_EXCEL):
        create_sample_input(INPUT_EXCEL)
        print(f"已生成示例: {INPUT_EXCEL}，请修改后重新运行")
        raise SystemExit(0)

    cases = read_shear_cases(INPUT_EXCEL)
    if not cases:
        raise SystemExit(f"未从 {INPUT_SHEET} 读取到有效工况")

    main_hw = get_expert_hwnd()
    main_hw = ensure_expert_visible(main_hw)
    ctrl_map = build_control_map(main_hw)
    print(
        f"控件映射 OK: b={ctrl_map['edit_b']}, h={ctrl_map['edit_h']}, "
        f"d={ctrl_map['edit_d']}, shear={ctrl_map.get('chk_shear')}"
    )

    res = []
    for case in cases:
        print(
            f"计算 {case['ID']}（Excel {len(case['loads'])} 行 → Expert {len(case['loads'])} 组荷载）..."
        )
        res.append(run_one(case, main_hw))

    export_result_summary(res, OUTPUT_EXCEL)
    print(f"全部完成，共 {len(res)} 个构件")
