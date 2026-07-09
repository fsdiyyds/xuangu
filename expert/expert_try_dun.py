import win32gui
import win32con
import win32api
import win32process
import ctypes
import json
import sys
import time
import pandas as pd
import shutil
import os
import glob

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXPERT_M_DIR = os.path.join(SCRIPT_DIR, "expert_M")

# ===================== 配置 =====================
TARGET_WINDOW = "EXPERT RC - Combined axial load and bending"
INPUT_EXCEL = os.path.join(EXPERT_M_DIR, "input_NM.xlsx")
OUTPUT_EXCEL = os.path.join(EXPERT_M_DIR, "result_final.xlsx")
SLEEP_SHORT = 0.4
SLEEP_CALC = 1.3
SLEEP_NOTE = 1.5
# Expert 默认输出路径；Note 按钮生成的 rc_note.rtf 在此
EXPERT_RTF_PATH = os.path.join(
    os.path.expanduser("~"), "Documents", "Autodesk", "Output", "rc_note.rtf"
)
# 按计算编号另存 RTF 的目标文件夹（可改成任意路径）
OUTPUT_RTF_DIR = os.path.join(EXPERT_M_DIR, "notes")
# 荷载表坐标（相对 GXWND 左上角）；运行 expert_M/calibrate_load.py 写入 load_coords.json
LOAD_COORDS_FILE = os.path.join(EXPERT_M_DIR, "load_coords.json")
LOAD_ROW_INDEX_COL_X = 20
LOAD_TYPE_COL_X = 69
LOAD_N_COL_X = 171
LOAD_M_COL_X = 264
LOAD_FIRST_ROW_Y = 2
LOAD_ROW_HEIGHT = 19.8
LOAD_CLEAR_TIMES = 10


def apply_load_coords():
    """优先读取 load_coords.json（calibrate_load.py 生成）"""
    global LOAD_ROW_INDEX_COL_X, LOAD_TYPE_COL_X, LOAD_N_COL_X, LOAD_M_COL_X
    global LOAD_FIRST_ROW_Y, LOAD_ROW_HEIGHT
    if not os.path.isfile(LOAD_COORDS_FILE):
        return False
    with open(LOAD_COORDS_FILE, encoding="utf-8") as f:
        d = json.load(f)
    LOAD_ROW_INDEX_COL_X = int(d["LOAD_ROW_INDEX_COL_X"])
    LOAD_TYPE_COL_X = int(d["LOAD_TYPE_COL_X"])
    LOAD_N_COL_X = int(d.get("LOAD_N_COL_X", LOAD_TYPE_COL_X + 102))
    LOAD_M_COL_X = int(d.get("LOAD_M_COL_X", LOAD_TYPE_COL_X + 195))
    LOAD_FIRST_ROW_Y = int(d["LOAD_FIRST_ROW_Y"])
    LOAD_ROW_HEIGHT = int(d["LOAD_ROW_HEIGHT"])
    return True


apply_load_coords()

EDIT_CLASSES = {"Edit", "TEdit", "TMaskEdit"}
RICHEDIT_CLASSES = {"RichEdit", "RichEdit20A", "RichEdit20W", "RichEdit50W"}
BUTTON_CLASS = {"Button", "TButton"}
DIM_CTRL_IDS = {"edit_b": 1227, "edit_h": 1228, "edit_d1": 1232, "edit_d2": 1233}
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


def build_control_map(main_hw):
    main_hw = ensure_expert_visible(main_hw)
    rows, page_hw = get_page_rows(main_hw)
    grid_rect, grid_hwnd = parse_load_grid_rect(rows)
    edit_b, edit_h, edit_d1, edit_d2 = find_dimension_edits(rows, page_hw)

    btn_calc = find_button_hwnd(rows, "calculate", "c a l c u l a t e")
    btn_note = find_button_hwnd(rows, "note", "n o t e")
    rad_beam = find_button_hwnd(rows, "beam")
    rad_col = find_button_hwnd(rows, "column")
    chk_seismic = find_button_hwnd(rows, "seismic")

    missing = [n for n, h in [("Calculate", btn_calc), ("Note", btn_note),
                                ("beam", rad_beam), ("column", rad_col)] if not h]
    if missing:
        raise RuntimeError(f"未找到按钮: {missing}")

    return {
        "edit_b": edit_b, "edit_h": edit_h, "edit_d1": edit_d1, "edit_d2": edit_d2,
        "rad_beam": rad_beam, "rad_col": rad_col, "chk_seismic": chk_seismic,
        "btn_calc": btn_calc, "btn_note": btn_note,
        "grid_rect": grid_rect, "grid_hwnd": grid_hwnd,
        "results_log_hwnd": find_results_log_edit(rows, grid_rect),
    }


def click_control(hwnd):
    if not is_valid_hwnd(hwnd):
        raise RuntimeError("控件句柄无效")
    win32gui.SendMessage(hwnd, win32con.BM_CLICK, 0, 0)
    time.sleep(SLEEP_SHORT)


def set_edit_text(hwnd, text):
    if not is_valid_hwnd(hwnd):
        raise RuntimeError("Edit 句柄无效")
    win32gui.SendMessage(hwnd, win32con.WM_SETTEXT, 0, str(text))
    time.sleep(SLEEP_SHORT)


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


def load_cell_coords(grid_rect, row_index, col_index):
    """三列分别点击：0=Load type, 1=N, 2=M"""
    gl, gt, _, gb = grid_rect
    y = gt + LOAD_FIRST_ROW_Y + row_index * LOAD_ROW_HEIGHT
    if y > gb - 8:
        return None
    col_x = (LOAD_TYPE_COL_X, LOAD_N_COL_X, LOAD_M_COL_X)[col_index]
    return gl + col_x, y


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
        time.sleep(0.35)
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
    time.sleep(0.2)
    dismiss_rcba_dialog()


def write_cell_editor(grid_hwnd, grid_rect, text):
    edit = get_cell_editor(grid_hwnd, grid_rect)
    if not edit:
        return False
    win32gui.SendMessage(edit, win32con.WM_SETTEXT, 0, str(text))
    time.sleep(0.08)
    return win32gui.GetWindowText(edit).strip() == str(text).strip()


def type_in_grid_cell(main_hw, grid_hwnd, grid_rect, text):
    grid_send_key(main_hw, grid_hwnd, win32con.VK_F2)
    time.sleep(0.3)
    if write_cell_editor(grid_hwnd, grid_rect, text):
        return
    grid_send_key(main_hw, grid_hwnd, win32con.VK_DELETE)
    time.sleep(0.06)
    grid_send_text(main_hw, grid_hwnd, text)


def click_and_type_cell(main_hw, grid_rect, grid_hwnd, row_index, col_index, text, label):
    """点击 GXWND 单元格 → 键盘消息直接发给网格（不经过 Results 文本框）"""
    coords = load_cell_coords(grid_rect, row_index, col_index)
    if not coords:
        print(f"  ✗ 第 {row_index + 1} 行 {label} 坐标无效")
        return False
    print(f"  → 点击 {label} ({int(coords[0])}, {int(coords[1])})  输入: {text}")
    click_grid_at(main_hw, grid_hwnd, coords[0], coords[1])
    time.sleep(0.3)
    dismiss_rcba_dialog()

    if not str(text):
        return True

    if col_index in (1, 2):
        type_in_grid_cell(main_hw, grid_hwnd, grid_rect, text)
        print(f"    F2 + 网格内写入")
    else:
        grid_send_text(main_hw, grid_hwnd, text)

    commit_cell_edit(main_hw, grid_hwnd)
    return True


def fill_load_row(main_hw, grid_rect, grid_hwnd, load_type, n_val, m_val, row_index):
    """
    最原始输入：逐格点击白格 → 键入 → 移到右侧 N/M 格继续
    多行荷载时 row_index 递增，自动点到下一行。
    """
    lt = format_cell_value(load_type)
    n_str = format_load_number(n_val)
    m_str = format_load_number(m_val)

    print(f"  --- 第 {row_index + 1} 行荷载: {lt}  N={n_str}  M={m_str} ---")
    click_and_type_cell(main_hw, grid_rect, grid_hwnd, row_index, 0, lt, "Load type")
    click_and_type_cell(main_hw, grid_rect, grid_hwnd, row_index, 1, n_str, "N")
    click_and_type_cell(main_hw, grid_rect, grid_hwnd, row_index, 2, m_str, "M")
    print(f"  → 第 {row_index + 1} 行完成")


def verify_design_page(main_hw):
    get_design_page_hwnd(main_hw)


def ensure_checkbox_checked(hwnd):
    if not is_valid_hwnd(hwnd):
        return
    if win32gui.SendMessage(hwnd, win32con.BM_GETCHECK, 0, 0) != win32con.BST_CHECKED:
        win32gui.SendMessage(hwnd, win32con.BM_CLICK, 0, 0)
    time.sleep(SLEEP_SHORT)


def dismiss_rcba_dialog():
    dialogs = []
    win32gui.EnumWindows(
        lambda hw, arr: arr.append(hw) or True
        if win32gui.IsWindowVisible(hw) and win32gui.GetWindowText(hw) == "rcba"
        else True,
        dialogs,
    )
    for dlg in dialogs:
        win32gui.EnumChildWindows(
            dlg,
            lambda hw, _: win32gui.SendMessage(hw, win32con.BM_CLICK, 0, 0)
            if win32gui.GetWindowText(hw) in ("确定", "OK") else True,
            None,
        )


def clear_loads_quick(main_hw, grid_rect, grid_hwnd, times=LOAD_CLEAR_TIMES):
    """点第 1 行左侧行号 → Delete 清空整行，重复 times 次"""
    print(f"  清空荷载：点行号 1 + Delete × {times}")
    coords = row_index_coords(grid_rect, 0)
    if not coords:
        return
    for _ in range(times):
        dismiss_rcba_dialog()
        click_grid_at(main_hw, grid_hwnd, coords[0], coords[1])
        time.sleep(0.25)
        grid_send_key(main_hw, grid_hwnd, win32con.VK_DELETE)
        time.sleep(0.25)
        dismiss_rcba_dialog()


def fill_all_load_rows(main_hw, load_rows, grid_rect, grid_hwnd, results_log_hw=None):
    blocked = None
    if is_valid_hwnd(results_log_hw):
        blocked = results_log_hw
        win32gui.EnableWindow(blocked, False)
        time.sleep(0.1)
    try:
        with expert_keyboard_session(main_hw) as main_hw:
            dismiss_rcba_dialog()
            if not is_valid_hwnd(grid_hwnd):
                print("  警告: 未找到 GXWND 句柄，请确认在 Design 页")
            clear_loads_quick(main_hw, grid_rect, grid_hwnd)
            print("  → 清空完成，开始填入荷载...")
            time.sleep(0.5)
            for i, lr in enumerate(load_rows):
                print(
                    f"  荷载行 {i + 1}: {lr['load_type']}  "
                    f"N={format_load_number(lr['N_kN'])}  M={format_load_number(lr['M_kNm'])}"
                )
                fill_load_row(
                    main_hw, grid_rect, grid_hwnd,
                    lr["load_type"], lr["N_kN"], lr["M_kNm"], i,
                )
    finally:
        if blocked:
            win32gui.EnableWindow(blocked, True)
    verify_design_page(main_hw)


def close_rc_note_viewer():
    """关闭已打开的 rc_note.rtf 查看器，避免 Note 时文件被占用"""

    def close_hw(hw, _):
        if not win32gui.IsWindowVisible(hw):
            return True
        cls = win32gui.GetClassName(hw)
        title = win32gui.GetWindowText(hw).lower()
        if TARGET_WINDOW.lower() in title:
            return True
        if cls in ("WordPadClass", "OpusApp") or "rc_note" in title or "rc_note.rtf" in title:
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


def normalize_load_columns(df):
    col_map = {}
    for c in df.columns:
        cl = str(c).strip().lower()
        if cl in ("load_type", "loadtype"):
            col_map[c] = "load_type"
        elif cl in ("n_kn", "n"):
            col_map[c] = "N_kN"
        elif cl in ("m_knm", "m"):
            col_map[c] = "M_kNm"
    return df.rename(columns=col_map)


def run_one(row, load_rows, main_hw):
    main_hw = ensure_expert_visible(main_hw)
    ctrls = build_control_map(main_hw)
    grid_rect = ctrls["grid_rect"]
    grid_hwnd = ctrls.get("grid_hwnd")

    # 1. 先填荷载（此时 Design 页最稳定，坐标已缓存）
    fill_all_load_rows(
        main_hw, load_rows, grid_rect, grid_hwnd, ctrls.get("results_log_hwnd")
    )

    # 2. 截面尺寸（SendMessage，不移动鼠标）
    set_edit_text(ctrls["edit_b"], row["b_cm"])
    set_edit_text(ctrls["edit_h"], row["h_cm"])
    set_edit_text(ctrls["edit_d1"], row["d1_cm"])
    set_edit_text(ctrls["edit_d2"], row["d2_cm"])

    # 3. 梁/柱
    click_control(ctrls["rad_col"] if int(row["is_column"]) == 1 else ctrls["rad_beam"])

    # 4. 勾选 Seismic detailing
    if ctrls.get("chk_seismic"):
        ensure_checkbox_checked(ctrls["chk_seismic"])

    # 5. 计算
    click_control(ctrls["btn_calc"])
    time.sleep(SLEEP_CALC)

    # 6. Note → 复制 RTF → 关闭查看器
    close_rc_note_viewer()
    click_control(ctrls["btn_note"])
    time.sleep(SLEEP_NOTE)
    new_rtf = save_rtf(row["ID"])

    return {"ID": row["ID"], "b": row["b_cm"], "h": row["h_cm"],
            "loads": len(load_rows), "RTF": new_rtf}


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] in ("--calibrate", "--test-coords"):
        sys.path.insert(0, EXPERT_M_DIR)
        import calibrate_load
        if sys.argv[1] == "--calibrate":
            calibrate_load.run_calibrate()
        else:
            calibrate_load.run_test()
        raise SystemExit(0)

    print("=" * 50)
    print("运行前请：打开 Expert → Design 页 → 窗口不要最小化")
    print("运行期间请勿操作鼠标键盘")
    print(f"RTF 另存目录: {OUTPUT_RTF_DIR}")
    if os.path.isfile(LOAD_COORDS_FILE):
        print(f"荷载坐标: 已加载 {LOAD_COORDS_FILE}")
        print(f"  行号1 X={LOAD_ROW_INDEX_COL_X}  Type={LOAD_TYPE_COL_X}  "
              f"N={LOAD_N_COL_X}  M={LOAD_M_COL_X}  首行Y={LOAD_FIRST_ROW_Y}  行高={LOAD_ROW_HEIGHT}")
    else:
        print("荷载坐标: 未校准！请先运行  cd expert_M && python calibrate_load.py")
        print("          或  python expert_try.py --calibrate")
    print("=" * 50)

    main_hw = get_expert_hwnd()
    main_hw = ensure_expert_visible(main_hw)
    ctrl_map = build_control_map(main_hw)
    print(f"控件映射 OK: b={ctrl_map['edit_b']}, h={ctrl_map['edit_h']}, "
          f"seismic={ctrl_map.get('chk_seismic')}")
    print(f"  荷载表 GXWND={ctrl_map.get('grid_hwnd')}  "
          f"Results框={ctrl_map.get('results_log_hwnd')}（填表时会临时禁用）")

    if not os.path.exists(INPUT_EXCEL):
        pd.DataFrame([{
            "ID": "C01", "b_cm": 25, "h_cm": 50, "d1_cm": 5, "d2_cm": 5,
            "is_column": 0, "load_type": "ULS", "N_kN": 1350, "M_kNm": 88,
        }]).to_excel(INPUT_EXCEL, index=False)
        print(f"已生成示例: {INPUT_EXCEL}，请修改后重新运行")
        raise SystemExit(0)

    df = pd.read_excel(INPUT_EXCEL)
    df = normalize_load_columns(df)
    if "load_type" not in df.columns:
        df["load_type"] = "ULS"
    if "ID" not in df.columns:
        df["ID"] = range(1, len(df) + 1)

    for col in ["b_cm", "h_cm", "d1_cm", "d2_cm"]:
        if col not in df.columns:
            raise RuntimeError(f"Excel 缺少列: {col}")
    if "is_column" not in df.columns:
        df["is_column"] = 0

    res = []
    for gid, group in df.groupby("ID", sort=False):
        loads = [{
            "load_type": str(r.get("load_type", "ULS")).strip(),
            "N_kN": float(r["N_kN"]),
            "M_kNm": float(r["M_kNm"]),
        } for _, r in group.iterrows()]
        print(f"计算构件 {gid}（{len(loads)} 组荷载）...")
        res.append(run_one(group.iloc[0], loads, main_hw))

    pd.DataFrame(res).to_excel(OUTPUT_EXCEL, index=False)
    print(f"完成 → {OUTPUT_EXCEL}")
