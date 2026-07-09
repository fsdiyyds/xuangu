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
import re
import threading

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ===================== 配置 =====================
TARGET_WINDOW = "EXPERT RC - Combined axial load and biaxial bending"
INPUT_EXCEL = os.path.join(SCRIPT_DIR, "input_NM_biaxial.xlsx")
OUTPUT_EXCEL = os.path.join(SCRIPT_DIR, "result_final_biaxial.xlsx")

# ---------- 等待时间（调大=更慢更稳，调小=更快但易出错）----------
SLEEP_SHORT = 0.05         # 控件点击、勾选 Seismic 等
SLEEP_CALC = 0.1          # Calculate 后额外稳定等待（弹框关闭后再等）
SLEEP_CALC_MAX = 5.0      # Calculate 后最长等待（轮询 rcba 直至关闭或超时）
SLEEP_NOTE = 1.5           # Note 后等待 RTF 生成
SLEEP_CELL_CLICK = 0.1    # 荷载表物理/模拟点击单元格后
SLEEP_CELL_SCROLL = 0.1    # Down 翻页后
SLEEP_COMMIT_CELL = 0.1    # 单格 Enter 提交后
SLEEP_TYPE_CELL = 0.01      # F2 进入单元格编辑后
SLEEP_GRID_CLEAR = 0.05    # 清空荷载（点行号 / Delete 每步）
SLEEP_LOAD_START = 0.05     # 清空荷载完毕、开始填入前
SLEEP_RCBA_POLL = 0.15     # 轮询 rcba 对话框间隔
SLEEP_KEYBOARD_SESSION = 0.2  # 荷载填表前 AttachThreadInput 抢焦点后
SLEEP_KEY_CHAR = 0.025     # 逐字符 WM_CHAR 间隔（grid_send_text）
SLEEP_KEY_VK = 0.04        # 单键 SendInput 间隔（send_vk）
SLEEP_SAFE_CLICK = 0.15    # safe_click 物理鼠标点击后
SLEEP_FIND_RTF = 0.3       # find_rc_note_file 轮询间隔
# Expert 默认输出路径；Note 按钮生成的 rc_note.rtf 在此
EXPERT_RTF_PATH = os.path.join(
    os.path.expanduser("~"), "Documents", "Autodesk", "Output", "rc_note.rtf"
)
# 按计算编号另存 RTF 的目标文件夹（可改成任意路径）
OUTPUT_RTF_DIR = os.path.join(SCRIPT_DIR, "notes")

# 荷载表坐标（相对 GXWND 左上角）；运行 calibrate_load.py 自动写入 load_coords.json
LOAD_COORDS_FILE = os.path.join(SCRIPT_DIR, "load_coords_biaxial.json")
LOAD_ROW_INDEX_COL_X = 20
LOAD_TYPE_COL_X = 69
LOAD_N_COL_X = 171
LOAD_MY_COL_X = 220
LOAD_MZ_COL_X = 281
LOAD_EXPAND_BTN_X = 350
LOAD_EXPAND_BTN_Y = 120
LOAD_FIRST_ROW_Y = 2
LOAD_ROW_HEIGHT = 19.8
LOAD_CLEAR_TIMES = 10
# 荷载表可见 9 行；第 10 行起先点第 9 行 Load type + Down 翻页，填表 Y 取第 8 行
LOAD_ROWS_BEFORE_EXPAND = 9
LOAD_SCROLL_CLICK_ROW = 8   # 0-based，翻页时物理点击第 9 行 Load type
LOAD_OVERFLOW_FILL_ROW = 7  # 0-based，第 10 行及以后填表 Y（第 8 行；第 9 行 Y 会点到滚动条）


def apply_load_coords():
    """优先读取 load_coords_biaxial.json（calibrate_load_biaxial.py 生成）"""
    global LOAD_ROW_INDEX_COL_X, LOAD_TYPE_COL_X, LOAD_N_COL_X
    global LOAD_MY_COL_X, LOAD_MZ_COL_X, LOAD_EXPAND_BTN_X, LOAD_EXPAND_BTN_Y
    global LOAD_FIRST_ROW_Y, LOAD_ROW_HEIGHT
    if not os.path.isfile(LOAD_COORDS_FILE):
        return False
    with open(LOAD_COORDS_FILE, encoding="utf-8") as f:
        d = json.load(f)
    LOAD_ROW_INDEX_COL_X = int(d["LOAD_ROW_INDEX_COL_X"])
    LOAD_TYPE_COL_X = int(d["LOAD_TYPE_COL_X"])
    LOAD_FIRST_ROW_Y = int(d["LOAD_FIRST_ROW_Y"])
    LOAD_ROW_HEIGHT = int(d["LOAD_ROW_HEIGHT"])
    LOAD_N_COL_X = int(d.get("LOAD_N_COL_X", LOAD_TYPE_COL_X + 102))
    LOAD_MY_COL_X = int(d.get("LOAD_MY_COL_X", LOAD_N_COL_X + 55))
    LOAD_MZ_COL_X = int(d.get("LOAD_MZ_COL_X", LOAD_MY_COL_X + 55))
    LOAD_EXPAND_BTN_X = int(d.get("LOAD_EXPAND_BTN_X", LOAD_MZ_COL_X + 40))
    LOAD_EXPAND_BTN_Y = int(d.get("LOAD_EXPAND_BTN_Y", LOAD_FIRST_ROW_Y + 80))
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


def read_expert_results(main_hw, grid_rect, chk_seismic_hw=None):
    """Calculate 完成后读取 Results 区输出"""
    rows, _ = get_page_rows(main_hw)
    as1_hw, as2_hw = find_results_as_edits(rows, grid_rect)
    as1 = _read_edit_value(as1_hw)
    as2 = _read_edit_value(as2_hw)
    rho, rho_min, rho_max = _parse_reinforcement_statics(rows, grid_rect)
    seismic = None
    if chk_seismic_hw and is_valid_hwnd(chk_seismic_hw):
        seismic = (
            win32gui.SendMessage(chk_seismic_hw, win32con.BM_GETCHECK, 0, 0)
            == win32con.BST_CHECKED
        )
    out = {
        "As1_cm2": as1,
        "As2_cm2": as2,
        "rho_pct": rho,
        "rho_min_pct": rho_min,
        "rho_max_pct": rho_max,
        "seismic_detail": seismic,
    }
    print(
        f"  → Results: As1={as1}  As2={as2}  "
        f"ρ={rho}%  ρmin={rho_min}%  ρmax={rho_max}%  "
        f"Seismic={seismic}"
    )
    return out


def build_control_map(main_hw):
    main_hw = ensure_expert_visible(main_hw)
    rows, page_hw = get_page_rows(main_hw)
    grid_rect, grid_hwnd = parse_load_grid_rect(rows)
    edit_b, edit_h, edit_d1, edit_d2 = find_dimension_edits(rows, page_hw)

    btn_calc = find_button_hwnd(rows, "calculate", "c a l c u l a t e")
    btn_note = find_button_hwnd(rows, "note", "n o t e")
    chk_seismic = find_button_hwnd(rows, "seismic")

    missing = [n for n, h in [("Calculate", btn_calc), ("Note", btn_note)] if not h]
    if missing:
        raise RuntimeError(f"未找到按钮: {missing}")

    return {
        "edit_b": edit_b, "edit_h": edit_h, "edit_d1": edit_d1, "edit_d2": edit_d2,
        "chk_seismic": chk_seismic,
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


def effective_screen_row(logical_row_index: int) -> int:
    """第 10 行及以后逻辑行填表 Y 取第 8 行屏幕坐标（避免点到滚动条）"""
    if logical_row_index >= LOAD_ROWS_BEFORE_EXPAND:
        return LOAD_OVERFLOW_FILL_ROW
    return logical_row_index


def load_col_screen_x(grid_rect, col_index):
    """列 X 与逻辑行无关，各列始终用校准得到的固定偏移"""
    gl, _, _, _ = grid_rect
    col_x = (LOAD_TYPE_COL_X, LOAD_N_COL_X, LOAD_MY_COL_X, LOAD_MZ_COL_X)[col_index]
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


def load_cell_coords(grid_rect, logical_row_index, col_index):
    """
    四列：0=Load type, 1=N, 2=MY, 3=MZ。
    X 仅由列号决定；Y 由屏幕行决定。逻辑第 10 行及以后 Y 取第 8 行。
    """
    screen_row = effective_screen_row(logical_row_index)
    return load_cell_coords_at_screen_row(grid_rect, screen_row, col_index)


def scroll_load_grid_down(main_hw, grid_hwnd, grid_rect):
    """
    第 10 行及以后：物理点第 9 行 Load type → Down → 再点第 8 行 Load type 确认。
    翻页点击用第 9 行 Y；翻页后填表用第 8 行 Y（第 9 行 Y 会落到滚动条）。
    """
    scroll_coords = load_cell_coords_at_screen_row(
        grid_rect, LOAD_SCROLL_CLICK_ROW, 0,
    )
    fill_coords = load_cell_coords_at_screen_row(
        grid_rect, LOAD_OVERFLOW_FILL_ROW, 0,
    )
    if not scroll_coords or not fill_coords:
        print("  ✗ 无法定位翻页/填表锚点单元格（Load type）")
        return
    print(
        f"  → 翻页：物理点第 {LOAD_SCROLL_CLICK_ROW + 1} 行 Load type "
        f"({int(scroll_coords[0])}, {int(scroll_coords[1])}) + 向下键"
    )
    safe_click(main_hw, scroll_coords[0], scroll_coords[1])
    time.sleep(SLEEP_CELL_CLICK)
    focus_grid(main_hw, grid_hwnd)
    grid_send_key(main_hw, grid_hwnd, win32con.VK_DOWN)
    time.sleep(SLEEP_CELL_SCROLL)
    print(
        f"  → 翻页后物理点第 {LOAD_OVERFLOW_FILL_ROW + 1} 行 Load type 填表锚点: "
        f"({int(fill_coords[0])}, {int(fill_coords[1])})"
    )
    safe_click(main_hw, fill_coords[0], fill_coords[1])
    time.sleep(SLEEP_CELL_CLICK)
    focus_grid(main_hw, grid_hwnd)


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
    physical=False,
):
    """点击单元格后输入；physical=True 时用物理鼠标（翻页后填表）"""
    coords = load_cell_coords(grid_rect, logical_row_index, col_index)
    if not coords:
        print(f"  ✗ 第 {logical_row_index + 1} 行 {label} 坐标无效")
        return False
    screen_row = effective_screen_row(logical_row_index)
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
    commit_cell_edit(main_hw, grid_hwnd)
    return True


def fill_load_row_overflow(
    main_hw, grid_rect, grid_hwnd, lt, n_str, my_str, mz_str, row_index,
):
    """
    第 10 行及以后：第 9 行 Load type 物理点击 + Down 翻页；
    再物理点击锚点行各列填表（避免焦点停在 MZ 列导致错位）。
    """
    scroll_load_grid_down(main_hw, grid_hwnd, grid_rect)
    screen_row = effective_screen_row(row_index)
    print(
        f"  --- 第 {row_index + 1} 行荷载: {lt}  "
        f"N={n_str}  MY={my_str}  MZ={mz_str}  "
        f"(屏幕行 {screen_row + 1}，物理点击逐格填表) ---"
    )
    click_and_type_cell(
        main_hw, grid_rect, grid_hwnd, row_index, 0, lt, "Load type", physical=True,
    )
    click_and_type_cell(
        main_hw, grid_rect, grid_hwnd, row_index, 1, n_str, "N", physical=True,
    )
    click_and_type_cell(
        main_hw, grid_rect, grid_hwnd, row_index, 2, my_str, "MY", physical=True,
    )
    click_and_type_cell(
        main_hw, grid_rect, grid_hwnd, row_index, 3, mz_str, "MZ", physical=True,
    )


def fill_load_row(
    main_hw, grid_rect, grid_hwnd, load_type, n_val, my_val, mz_val, row_index,
):
    """
    第 1～9 行：逐格点击 Load type → N → MY → MZ。
    第 10 行及以后：第 9 行 Load type 物理点击翻页 + 物理点击逐格填表。
    """
    lt = format_cell_value(load_type)
    n_str = format_load_number(n_val)
    my_str = format_load_number(my_val)
    mz_str = format_load_number(mz_val)

    if row_index >= LOAD_ROWS_BEFORE_EXPAND:
        fill_load_row_overflow(
            main_hw, grid_rect, grid_hwnd, lt, n_str, my_str, mz_str, row_index,
        )
        print(f"  → 第 {row_index + 1} 行完成")
        return

    screen_row = effective_screen_row(row_index)
    print(
        f"  --- 第 {row_index + 1} 行荷载: {lt}  "
        f"N={n_str}  MY={my_str}  MZ={mz_str}  "
        f"(屏幕行 {screen_row + 1}) ---"
    )
    click_and_type_cell(main_hw, grid_rect, grid_hwnd, row_index, 0, lt, "Load type")
    click_and_type_cell(main_hw, grid_rect, grid_hwnd, row_index, 1, n_str, "N")
    click_and_type_cell(main_hw, grid_rect, grid_hwnd, row_index, 2, my_str, "MY")
    click_and_type_cell(main_hw, grid_rect, grid_hwnd, row_index, 3, mz_str, "MZ")
    print(f"  → 第 {row_index + 1} 行完成")


def verify_design_page(main_hw):
    get_design_page_hwnd(main_hw)


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


def fill_all_load_rows(main_hw, load_rows, grid_rect, grid_hwnd, results_log_hw=None):
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
            print("  → 清空完成，开始填入荷载...")
            time.sleep(SLEEP_LOAD_START)
            for i, lr in enumerate(load_rows):
                print(
                    f"  荷载行 {i + 1}: {lr['load_type']}  "
                    f"N={format_load_number(lr['N_kN'])}  "
                    f"MY={format_load_number(lr['My_kNm'])}  "
                    f"MZ={format_load_number(lr['Mz_kNm'])}"
                )
                fill_load_row(
                    main_hw, grid_rect, grid_hwnd,
                    lr["load_type"], lr["N_kN"], lr["My_kNm"], lr["Mz_kNm"], i,
                )
    finally:
        if blocked:
            win32gui.EnableWindow(blocked, True)
    verify_design_page(main_hw)
    print("  → 荷载表填写完成")


def _is_rc_note_window_title(title: str) -> bool:
    t = (title or "").strip().lower().replace("\\", "/")
    if not t:
        return False
    return "rc_note.rtf" in t or t == "rc_note" or t.startswith("rc_note ")


def close_rc_note_viewer():
    """仅关闭正在查看 rc_note.rtf 的窗口，不关其他 Word 文档。"""

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


def _map_input_column(col_name: str) -> str:
    """表头映射：N(kN)→N_kN，My(kNm)→My_kNm，Mz(kNm)→Mz_kNm，LOADTYPE→load_type"""
    key = _norm_header_key(col_name)
    exact = {
        "id": "ID",
        "bcm": "b_cm",
        "hcm": "h_cm",
        "d1cm": "d1_cm",
        "d2cm": "d2_cm",
        "loadtype": "load_type",
        "load_type": "load_type",
        "名称": "load_type",
        "n(kn)": "N_kN",
        "n_kn": "N_kN",
        "my(knm)": "My_kNm",
        "my_knm": "My_kNm",
        "mz(knm)": "Mz_kNm",
        "mz_knm": "Mz_kNm",
    }
    if key in exact:
        return exact[key]
    if key.startswith("my(") or key == "my":
        return "My_kNm"
    if key.startswith("mz(") or key == "mz":
        return "Mz_kNm"
    if key.startswith("n(") and "kn" in key and not key.startswith("ny"):
        return "N_kN"
    return str(col_name).strip()


def normalize_input_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename = {c: _map_input_column(c) for c in df.columns}
    return df.rename(columns=rename)


def forward_fill_merged_cells(df: pd.DataFrame) -> pd.DataFrame:
    """Excel 合并单元格读入后下方为 NaN，向下填充为合并格内的同一值"""
    out = df.copy()
    for col in out.columns:
        out[col] = out[col].ffill()
    return out


def _clean_load_type(val) -> str:
    if pd.isna(val):
        return "ULS"
    s = str(val).strip()
    if not s:
        return "ULS"
    # force_aggregate「名称」如 ULS MAX → Expert 取 ULS
    return s.split()[0].upper()


def read_input_excel(path: str) -> pd.DataFrame:
    """
    读取输入表：处理合并单元格，提取 N(kN)、My(kNm)、Mz(kNm)、LOADTYPE。
    Qz、Qy 等列忽略，不填入 Expert。
    """
    df = pd.read_excel(path, dtype=object)
    df = normalize_input_columns(df)
    df = forward_fill_merged_cells(df)

    if "ID" not in df.columns:
        df["ID"] = range(1, len(df) + 1)
    if "load_type" not in df.columns:
        df["load_type"] = "ULS"

    for col in ("N_kN", "My_kNm", "Mz_kNm"):
        if col not in df.columns:
            raise RuntimeError(
                f"Excel 缺少内力列 {col}，需要表头含 N(kN)、My(kNm)、Mz(kNm)"
            )
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["load_type"] = df["load_type"].map(_clean_load_type)

    # 去掉内力全空的行
    has_force = df[["N_kN", "My_kNm", "Mz_kNm"]].notna().any(axis=1)
    df = df.loc[has_force].copy()

    for col in ["b_cm", "h_cm", "d1_cm", "d2_cm"]:
        if col not in df.columns:
            raise RuntimeError(f"Excel 缺少列: {col}")

    return df


def run_one(row, load_rows, main_hw):
    main_hw = ensure_expert_visible(main_hw)
    ctrls = build_control_map(main_hw)
    grid_rect = ctrls["grid_rect"]
    grid_hwnd = ctrls.get("grid_hwnd")

    # 1. 先填荷载（此时 Design 页最稳定，坐标已缓存）
    fill_all_load_rows(
        main_hw, load_rows, grid_rect, grid_hwnd, ctrls.get("results_log_hwnd")
    )

    # 2. 截面尺寸（b/h 超限填写后可能弹出 rcba，处理逻辑一致）
    print("  → 填写截面尺寸 b, h, d1, d2 ...")
    set_dimension_edit(main_hw, ctrls["edit_b"], row["b_cm"], "填写 b 后")
    set_dimension_edit(main_hw, ctrls["edit_h"], row["h_cm"], "填写 h 后")
    set_edit_text(ctrls["edit_d1"], row["d1_cm"])
    set_edit_text(ctrls["edit_d2"], row["d2_cm"])

    # 3. 勾选 Seismic detailing（双轴弯矩柱界面无需选 beam/column）
    if ctrls.get("chk_seismic"):
        ensure_checkbox_checked(ctrls["chk_seismic"])

    # 4. 计算（PostMessage 异步点击，再轮询 rcba 点「是」；SendMessage 会死锁）
    post_click_control(ctrls["btn_calc"], "Calculate")
    time.sleep(0.35)
    if not wait_after_calculate(main_hw):
        raise RuntimeError("rcba 几何警告框未关闭，已中止（请检查 b/h 几何限制提示）")

    result_outputs = read_expert_results(
        main_hw, grid_rect, ctrls.get("chk_seismic"),
    )

    # 5. Note → 复制 RTF → 关闭查看器（须确认无 rcba 遮挡）
    if not _all_rcba_closed(main_hw):
        raise RuntimeError("rcba 警告框仍打开，无法继续 Note")
    close_rc_note_viewer()
    post_click_control(ctrls["btn_note"], "Note")
    time.sleep(SLEEP_NOTE)
    new_rtf = save_rtf(row["ID"])

    return {
        "ID": row["ID"],
        "b": row["b_cm"],
        "h": row["h_cm"],
        "loads": len(load_rows),
        **result_outputs,
        "RTF": new_rtf,
    }


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] in ("--calibrate", "--test-coords"):
        import calibrate_load_biaxial as calibrate_load
        if sys.argv[1] == "--calibrate":
            calibrate_load.run_calibrate()
        else:
            calibrate_load.run_test()
        raise SystemExit(0)

    print("=" * 50)
    print("Expert 双轴弯矩版 (biaxial bending)")
    print("运行前请：打开 Expert → Design 页 → 窗口不要最小化")
    print("运行期间请勿操作鼠标键盘")
    print(f"RTF 另存目录: {OUTPUT_RTF_DIR}")
    if os.path.isfile(LOAD_COORDS_FILE):
        print(f"荷载坐标: 已加载 {LOAD_COORDS_FILE}")
        print(
            f"  行号1 X={LOAD_ROW_INDEX_COL_X}  Type={LOAD_TYPE_COL_X}  "
            f"N={LOAD_N_COL_X}  MY={LOAD_MY_COL_X}  MZ={LOAD_MZ_COL_X}  "
            f"翻页点击行={LOAD_SCROLL_CLICK_ROW + 1}  "
            f"翻页填表Y行={LOAD_OVERFLOW_FILL_ROW + 1}  "
            f"首行Y={LOAD_FIRST_ROW_Y}  行高={LOAD_ROW_HEIGHT}"
        )
    else:
        print("荷载坐标: 未校准！请先运行  python expert_try_biaxial.py --calibrate")
        print("          或测试坐标  python expert_try_biaxial.py --test-coords")
    print("=" * 50)

    main_hw = get_expert_hwnd()
    main_hw = ensure_expert_visible(main_hw)
    ctrl_map = build_control_map(main_hw)
    print(f"控件映射 OK: b={ctrl_map['edit_b']}, h={ctrl_map['edit_h']}, "
          f"seismic={ctrl_map.get('chk_seismic')}")
    print(f"  荷载表 GXWND={ctrl_map.get('grid_hwnd')}  "
          f"Results框={ctrl_map.get('results_log_hwnd')}（填表时会临时禁用）")

    if not os.path.exists(INPUT_EXCEL):
        sample_rows = []
        for lt, n, my, mz in [
            ("ULS", 1161.14, 2919.35, -674.35),
            ("ULS", 1053.35, 2484.03, -12.95),
            ("ALS", 900.0, 2000.0, -500.0),
            ("SLS", 1084.19, 2114.47, -595.65),
        ]:
            sample_rows.append({
                "ID": "C0B-B 1#", "b_cm": 100, "h_cm": 250,
                "d1_cm": 6, "d2_cm": 6, "LOADTYPE": lt,
                "N(kN)": n, "My(kNm)": my, "Mz(kNm)": mz,
            })
        pd.DataFrame(sample_rows).to_excel(INPUT_EXCEL, index=False)
        print(f"已生成示例: {INPUT_EXCEL}，请修改后重新运行")
        raise SystemExit(0)

    df = read_input_excel(INPUT_EXCEL)
    print(f"已读取输入表: {len(df)} 行荷载（合并格已向下填充）")

    res = []
    for gid, group in df.groupby("ID", sort=False):
        loads = [{
            "load_type": str(r.get("load_type", "ULS")).strip(),
            "N_kN": float(r["N_kN"]),
            "My_kNm": float(r["My_kNm"]),
            "Mz_kNm": float(r["Mz_kNm"]),
        } for _, r in group.iterrows()]
        print(f"计算构件 {gid}（{len(loads)} 组荷载）...")
        res.append(run_one(group.iloc[0], loads, main_hw))

    pd.DataFrame(res).to_excel(OUTPUT_EXCEL, index=False)
    print(f"完成 → {OUTPUT_EXCEL}")
