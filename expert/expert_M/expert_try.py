"""
Expert RC — 轴力+弯矩 (N-M) 验算自动化

用法：
  cd expert/expert_M
  python expert_try.py
  python expert_try.py input_NM_纵向.xlsx input_NM_横向.xlsx   # 按顺序批量计算
  python expert_try.py --output-base results a.xlsx b.xlsx     # 输出到 results/{文件名}/
  python expert_try.py --calibrate    # 校准荷载表坐标
  python extract_rtf_results.py       # 从 notes/ 提取结果

输入：一个或多个 input_NM.xlsx
输出：每个 xlsx 对应文件夹 {output_base}/{xlsx文件名}/
      ├── result_final.xlsx（配筋汇总）
      └── notes/{ID}.rtf
"""
import argparse
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
from typing import Optional

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ===================== 配置 =====================
TARGET_WINDOW = "EXPERT RC - Combined axial load and bending"
INPUT_EXCEL = os.path.join(SCRIPT_DIR, "input_NM_纵向.xlsx")
# 默认单文件时的 RTF 目录（批量模式按各 xlsx 独立子文件夹）
OUTPUT_RTF_DIR = os.path.join(SCRIPT_DIR, "notes")

RESULT_COLUMNS = [
    "ID", "distance_m", "b_cm", "h_cm", "d1_cm", "d2_cm", "loads",
    "As1_cm2", "As2_cm2", "As_min_cm2", "As_max_cm2", "As_design_cm2",
    "rho_pct", "rho_min_pct", "rho_max_pct", "RTF",
]

# ---------- 等待时间 ----------
SLEEP_SHORT = 0.05
SLEEP_CALC = 0.1
SLEEP_CALC_MAX = 5.0
SLEEP_NOTE = 1.5
SLEEP_CELL_CLICK = 0.1
SLEEP_CELL_SCROLL = 0.1
SLEEP_COMMIT_CELL = 0.1
SLEEP_TYPE_CELL = 0.01
SLEEP_GRID_CLEAR = 0.05
SLEEP_LOAD_START = 0.05
SLEEP_RCBA_POLL = 0.15
# Expert 默认输出路径；Note 按钮生成的 rc_note.rtf 在此
EXPERT_RTF_PATH = os.path.join(
    os.path.expanduser("~"), "Documents", "Autodesk", "Output", "rc_note.rtf"
)

# 荷载表坐标
LOAD_COORDS_FILE = os.path.join(SCRIPT_DIR, "load_coords.json")
LOAD_ROW_INDEX_COL_X = 20
LOAD_TYPE_COL_X = 69
LOAD_N_COL_X = 171
LOAD_M_COL_X = 264
LOAD_FIRST_ROW_Y = 2
LOAD_ROW_HEIGHT = 19.8
LOAD_CLEAR_TIMES = 10
# 荷载表可见 9 行；第 10 行起先点第 9 行 Load type + Down 翻页，填表 Y 取第 8 行
LOAD_ROWS_BEFORE_EXPAND = 9
LOAD_SCROLL_CLICK_ROW = 8   # 0-based，翻页时物理点击第 9 行 Load type
LOAD_OVERFLOW_FILL_ROW = 7  # 0-based，第 10 行及以后填表 Y（第 8 行）


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
    """Calculate 完成后读取 Results 区配筋输出（与界面一致）"""
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


def parse_reinforcement_from_rtf(rtf_path):
    """从计算书 RTF「Sections d'Acier」段落提取配筋。"""
    if not rtf_path or not os.path.isfile(rtf_path):
        return {}
    from extract_rtf_results import extract_rtf_fields, read_rtf_text
    from pathlib import Path as _Path
    fields = extract_rtf_fields(read_rtf_text(_Path(rtf_path)))
    return {k: v for k, v in fields.items() if k.startswith(("As", "rho"))}


def _merge_reinforcement(ui: dict, rtf: dict) -> dict:
    """RTF 优先（含 As_min/max）；界面 Results 作补充。"""
    keys = (
        "As1_cm2", "As2_cm2", "As_min_cm2", "As_max_cm2",
        "rho_pct", "rho_min_pct", "rho_max_pct",
    )
    out = {}
    for k in keys:
        rv, uv = rtf.get(k), ui.get(k)
        out[k] = rv if rv is not None else uv
    return out


def _governing_as(reinf: dict):
    """控制配筋面积：max(As1, As2, As_min)。"""
    vals = [
        reinf.get("As1_cm2"), reinf.get("As2_cm2"), reinf.get("As_min_cm2"),
    ]
    nums = [float(v) for v in vals if v is not None and not pd.isna(v)]
    return max(nums) if nums else None


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
    """同步点击（勿用于会弹出 rcba 模态框的操作，会死锁）"""
    if not is_valid_hwnd(hwnd):
        raise RuntimeError("控件句柄无效")
    win32gui.SendMessage(hwnd, win32con.BM_CLICK, 0, 0)
    time.sleep(SLEEP_SHORT)


def post_click_control(hwnd, label=""):
    """异步点击；Calculate/Note 会弹 rcba 时必须用 PostMessage。"""
    if not is_valid_hwnd(hwnd):
        raise RuntimeError("控件句柄无效")
    if label:
        print(f"  → 点击: {label}")
    win32gui.PostMessage(hwnd, win32con.BM_CLICK, 0, 0)
    time.sleep(SLEEP_SHORT)


def set_edit_text(hwnd, text, main_hw=None, dismiss_rcba_label=None):
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
    set_edit_text(hwnd, value, main_hw=main_hw, dismiss_rcba_label=label)
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
    """第 10 行及以后逻辑行填表 Y 取第 8 行屏幕坐标"""
    if logical_row_index >= LOAD_ROWS_BEFORE_EXPAND:
        return LOAD_OVERFLOW_FILL_ROW
    return logical_row_index


def load_col_screen_x(grid_rect, col_index):
    gl, _, _, _ = grid_rect
    col_x = (LOAD_TYPE_COL_X, LOAD_N_COL_X, LOAD_M_COL_X)[col_index]
    return gl + col_x


def load_row_screen_y(grid_rect, screen_row_index):
    gl, gt, _, gb = grid_rect
    y = gt + LOAD_FIRST_ROW_Y + screen_row_index * LOAD_ROW_HEIGHT
    if y > gb - 8:
        return None
    return y


def load_cell_coords_at_screen_row(grid_rect, screen_row_index, col_index):
    y = load_row_screen_y(grid_rect, screen_row_index)
    if y is None:
        return None
    return load_col_screen_x(grid_rect, col_index), y


def load_cell_coords(grid_rect, logical_row_index, col_index):
    """三列：0=Load type, 1=N, 2=M"""
    screen_row = effective_screen_row(logical_row_index)
    return load_cell_coords_at_screen_row(grid_rect, screen_row, col_index)


def scroll_load_grid_down(main_hw, grid_hwnd, grid_rect):
    """
    第 10 行及以后：物理点第 9 行 Load type → Down → 再点第 8 行 Load type 确认。
    """
    scroll_coords = load_cell_coords_at_screen_row(grid_rect, LOAD_SCROLL_CLICK_ROW, 0)
    fill_coords = load_cell_coords_at_screen_row(grid_rect, LOAD_OVERFLOW_FILL_ROW, 0)
    if not scroll_coords or not fill_coords:
        print("  ✗ 无法定位翻页/填表锚点（Load type）")
        return False
    main_hw = ensure_expert_visible(main_hw)
    print(
        f"  → 翻页：物理点第 {LOAD_SCROLL_CLICK_ROW + 1} 行 Load type "
        f"({int(scroll_coords[0])}, {int(scroll_coords[1])}) + Down"
    )
    safe_click(main_hw, scroll_coords[0], scroll_coords[1])
    time.sleep(SLEEP_CELL_CLICK)
    focus_grid(main_hw, grid_hwnd)
    grid_send_key(main_hw, grid_hwnd, win32con.VK_DOWN)
    time.sleep(SLEEP_CELL_SCROLL)
    print(
        f"  → 翻页后填表锚点：物理点第 {LOAD_OVERFLOW_FILL_ROW + 1} 行 Load type "
        f"({int(fill_coords[0])}, {int(fill_coords[1])})"
    )
    safe_click(main_hw, fill_coords[0], fill_coords[1])
    time.sleep(SLEEP_CELL_CLICK)
    focus_grid(main_hw, grid_hwnd)
    return True


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
    time.sleep(SLEEP_COMMIT_CELL)


def commit_cell_in_place(main_hw, grid_rect, grid_hwnd, logical_row_index):
    """提交当前格但不 Enter 换行（避免 M 列 Enter 误生成 ELU 0/0 空行）"""
    screen_row = effective_screen_row(logical_row_index)
    coords = load_cell_coords_at_screen_row(grid_rect, screen_row, 0)
    if coords and is_valid_hwnd(grid_hwnd):
        click_grid_at(main_hw, grid_hwnd, coords[0], coords[1])
    else:
        grid_send_key(main_hw, grid_hwnd, win32con.VK_TAB)
    time.sleep(SLEEP_COMMIT_CELL)


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
    type_in_grid_cell(main_hw, grid_hwnd, grid_rect, text)
    print(f"    F2 + 网格内写入")


def click_and_type_cell(
    main_hw, grid_rect, grid_hwnd, logical_row_index, col_index, text, label,
    physical=False, commit_in_place=False,
):
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

    if not str(text):
        return True

    _type_into_selected_cell(main_hw, grid_hwnd, grid_rect, col_index, text)
    if commit_in_place:
        commit_cell_in_place(main_hw, grid_rect, grid_hwnd, logical_row_index)
    else:
        commit_cell_edit(main_hw, grid_hwnd)
    return True


def fill_load_row_overflow(
    main_hw, grid_rect, grid_hwnd, lt, n_str, m_str, row_index,
):
    if not scroll_load_grid_down(main_hw, grid_hwnd, grid_rect):
        raise RuntimeError(f"第 {row_index + 1} 行荷载翻页失败，已中止")
    screen_row = effective_screen_row(row_index)
    print(
        f"  --- 第 {row_index + 1} 行荷载: {lt}  N={n_str}  M={m_str}  "
        f"(屏幕行 {screen_row + 1}，物理点击逐格填表) ---"
    )
    for col, text, label in (
        (0, lt, "Load type"),
        (1, n_str, "N"),
        (2, m_str, "M"),
    ):
        click_and_type_cell(
            main_hw, grid_rect, grid_hwnd, row_index, col, text, label,
            physical=True, commit_in_place=True,
        )


def fill_load_row(
    main_hw, grid_rect, grid_hwnd, load_type, n_val, m_val, row_index, total_rows,
):
    """
    第 1～9 行：逐格点击 Load type → N → M。
    第 10 行及以后：第 9 行 Load type 物理点击翻页 + 物理点击逐格填表。
    M 列用 commit_in_place，避免 Enter 在下一行自动生成 ELU 0/0 空行。
    """
    lt = format_cell_value(load_type)
    n_str = format_load_number(n_val)
    m_str = format_load_number(m_val)
    is_last_row = row_index >= total_rows - 1
    is_bottom_visible = row_index == LOAD_ROWS_BEFORE_EXPAND - 1

    if row_index >= LOAD_ROWS_BEFORE_EXPAND:
        fill_load_row_overflow(main_hw, grid_rect, grid_hwnd, lt, n_str, m_str, row_index)
        print(f"  → 第 {row_index + 1} 行完成")
        return

    screen_row = effective_screen_row(row_index)
    print(
        f"  --- 第 {row_index + 1} 行荷载: {lt}  N={n_str}  M={m_str}  "
        f"(屏幕行 {screen_row + 1}) ---"
    )
    for col, text, label in (
        (0, lt, "Load type"),
        (1, n_str, "N"),
        (2, m_str, "M"),
    ):
        in_place = is_last_row or is_bottom_visible or col == 2
        click_and_type_cell(
            main_hw, grid_rect, grid_hwnd, row_index, col, text, label,
            commit_in_place=in_place,
        )
    print(f"  → 第 {row_index + 1} 行完成")


def verify_design_page(main_hw):
    get_design_page_hwnd(main_hw)


def ensure_checkbox_checked(hwnd):
    if not is_valid_hwnd(hwnd):
        return
    if win32gui.SendMessage(hwnd, win32con.BM_GETCHECK, 0, 0) != win32con.BST_CHECKED:
        win32gui.SendMessage(hwnd, win32con.BM_CLICK, 0, 0)
    time.sleep(SLEEP_SHORT)


def _is_rcba_yes_button(text):
    t = (text or "").strip()
    if not t:
        return False
    if t in ("确定", "OK", "Yes", "是", "是(Y)", "是 (Y)", "是（Y）", "是(&Y)"):
        return True
    return t.startswith("是") and "否" not in t


_RCBA_HINT_SNIPPETS = (
    "do you wish to continue", "wish to continue", "continue calculations",
    "condition is not fulfilled", "not fulfilled", "may be incorrect",
    "currently set values", "warning", "exceed", "maximum", "allowable",
    "geometrical data", "geometrical", "geometry", "200 cm", "b <=", "h <=",
    "loading data", "load data", "moment", "bending moment", "bending",
    "axial load", "axial force", "n-m", "interaction diagram", "interaction",
    "kn·m", "knm", "继续", "是否继续", "几何", "不正确", "弯矩", "轴力",
    "荷载", "超出", "警告",
)


def _static_text_is_rcba_hint(text: str) -> bool:
    t = (text or "").lower()
    if not t.strip():
        return False
    if any(snippet in t for snippet in _RCBA_HINT_SNIPPETS):
        return True
    if re.search(r"\b(n|m)\s*[=:]", t):
        return True
    if re.search(r"\d+\.?\d*\s*(kn|knm|kn·m|kn\*m)", t):
        return True
    return False


def _is_rcba_warning_dialog(dlg):
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
    _set_foreground_window(dlg)
    time.sleep(0.1)


def _enum_child_buttons(parent_hw):
    buttons = []

    def walk(hw, _):
        if win32gui.GetClassName(hw) in BUTTON_CLASS:
            buttons.append(hw)
        return True

    win32gui.EnumChildWindows(parent_hw, walk, None)
    return buttons


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
    deadline = time.time() + timeout
    closed = 0
    while time.time() < deadline:
        dialogs = _find_rcba_dialogs(main_hw)
        if not dialogs:
            if closed:
                print(f"  → 共关闭 {closed} 个 rcba 警告框，继续后续步骤")
            elif label:
                print("  → rcba 警告框已关闭，继续后续步骤")
            time.sleep(SLEEP_SHORT)
            return True

        for dlg in dialogs:
            if not is_valid_hwnd(dlg):
                continue
            print(
                f"  → 发现 rcba 警告框（{label or '自动处理'}）"
                f" hwnd={dlg}，点击「是」..."
            )
            if _click_rcba_yes(dlg, main_hw):
                closed += 1
            time.sleep(0.12)

        if _all_rcba_closed(main_hw):
            if closed:
                print(f"  → 共关闭 {closed} 个 rcba 警告框，继续后续步骤")
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
    if not _find_rcba_dialogs(main_hw):
        return False
    return ensure_rcba_dismissed(main_hw, timeout=3.0, label="")


def wait_after_calculate(main_hw):
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
            print(
                f"  → 清空完成，开始填入荷载..."
                f"（界面可见约 {LOAD_ROWS_BEFORE_EXPAND} 行，"
                f"第 {LOAD_ROWS_BEFORE_EXPAND + 1} 行起翻页；"
                f"翻页点第 {LOAD_SCROLL_CLICK_ROW + 1} 行，"
                f"填表 Y 第 {LOAD_OVERFLOW_FILL_ROW + 1} 行）"
            )
            time.sleep(SLEEP_LOAD_START)
            total = len(load_rows)
            for i, lr in enumerate(load_rows):
                print(
                    f"  荷载行 {i + 1}/{total}: {lr['load_type']}  "
                    f"N={format_load_number(lr['N_kN'])}  M={format_load_number(lr['M_kNm'])}"
                )
                fill_load_row(
                    main_hw, grid_rect, grid_hwnd,
                    lr["load_type"], lr["N_kN"], lr["M_kNm"], i, total,
                )
    finally:
        if blocked:
            win32gui.EnableWindow(blocked, True)
    verify_design_page(main_hw)
    print(f"  → 荷载表填写完成（共 {len(load_rows)} 组）")


def _is_rc_note_window_title(title: str) -> bool:
    t = (title or "").strip().lower().replace("\\", "/")
    if not t:
        return False
    return "rc_note.rtf" in t or t == "rc_note" or t.startswith("rc_note ")


def close_rc_note_viewer():
    """仅关闭正在查看 rc_note.rtf 的窗口，避免 Note 时文件被占用"""

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


def save_rtf(note_id, rtf_dir=None):
    """从 Expert 输出目录复制 rc_note.rtf，按编号另存到 rtf_dir"""
    rtf_dir = rtf_dir or OUTPUT_RTF_DIR
    os.makedirs(rtf_dir, exist_ok=True)
    safe_id = str(note_id).strip().replace(os.sep, "_")
    dest = os.path.join(rtf_dir, f"{safe_id}.rtf")

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


def run_one(row, load_rows, main_hw, rtf_dir=None):
    main_hw = ensure_expert_visible(main_hw)
    ctrls = build_control_map(main_hw)
    grid_rect = ctrls["grid_rect"]
    grid_hwnd = ctrls.get("grid_hwnd")

    fill_all_load_rows(
        main_hw, load_rows, grid_rect, grid_hwnd, ctrls.get("results_log_hwnd")
    )

    print("  → 填写截面尺寸 b, h, d1, d2 ...")
    set_dimension_edit(main_hw, ctrls["edit_b"], row["b_cm"], "填写 b 后")
    set_dimension_edit(main_hw, ctrls["edit_h"], row["h_cm"], "填写 h 后")
    set_edit_text(ctrls["edit_d1"], row["d1_cm"])
    set_edit_text(ctrls["edit_d2"], row["d2_cm"])

    click_control(ctrls["rad_col"] if int(row["is_column"]) == 1 else ctrls["rad_beam"])

    if ctrls.get("chk_seismic"):
        ensure_checkbox_checked(ctrls["chk_seismic"])

    post_click_control(ctrls["btn_calc"], "Calculate")
    time.sleep(0.35)
    if not wait_after_calculate(main_hw):
        raise RuntimeError("rcba 警告框未关闭，已中止")

    ui_results = read_expert_results(
        main_hw, grid_rect, ctrls.get("chk_seismic"),
    )

    if not _all_rcba_closed(main_hw):
        raise RuntimeError("rcba 警告框仍打开，无法继续 Note")
    close_rc_note_viewer()
    post_click_control(ctrls["btn_note"], "Note")
    time.sleep(SLEEP_NOTE)
    new_rtf = save_rtf(row["ID"], rtf_dir=rtf_dir)

    rtf_results = parse_reinforcement_from_rtf(new_rtf)
    reinf = _merge_reinforcement(ui_results, rtf_results)
    as_design = _governing_as(reinf)
    if rtf_results:
        print(
            f"  → RTF 配筋: As1={reinf.get('As1_cm2')}  As2={reinf.get('As2_cm2')}  "
            f"As_min={reinf.get('As_min_cm2')}  As_max={reinf.get('As_max_cm2')}  "
            f"As_design={as_design}"
        )

    out = {
        "ID": row["ID"],
        "distance_m": row.get("distance_m"),
        "b_cm": row["b_cm"],
        "h_cm": row["h_cm"],
        "d1_cm": row["d1_cm"],
        "d2_cm": row["d2_cm"],
        "loads": len(load_rows),
        **reinf,
        "As_design_cm2": as_design,
        "RTF": new_rtf,
    }
    return out


def resolve_job_paths(excel_path, output_base=None):
    """每个 input xlsx → 独立输出目录（以 xlsx 文件名命名）。"""
    output_base = output_base or SCRIPT_DIR
    stem = os.path.splitext(os.path.basename(excel_path))[0]
    job_dir = os.path.join(output_base, stem)
    rtf_dir = os.path.join(job_dir, "notes")
    result_path = os.path.join(job_dir, "result_final.xlsx")
    return job_dir, rtf_dir, result_path


def process_input_excel(excel_path, main_hw, rtf_dir, result_path):
    """读取单个 input_NM.xlsx，顺序计算并写入 result_path。"""
    df = pd.read_excel(excel_path)
    df = normalize_load_columns(df)
    if "load_type" not in df.columns:
        df["load_type"] = "ULS"
    if "ID" not in df.columns:
        df["ID"] = range(1, len(df) + 1)

    for col in ["b_cm", "h_cm", "d1_cm", "d2_cm"]:
        if col not in df.columns:
            raise RuntimeError(f"{excel_path} 缺少列: {col}")
    if "is_column" not in df.columns:
        df["is_column"] = 0

    os.makedirs(rtf_dir, exist_ok=True)
    os.makedirs(os.path.dirname(result_path), exist_ok=True)

    res = []
    for gid, group in df.groupby("ID", sort=False):
        loads = []
        for _, r in group.iterrows():
            n_kn = float(r["N_kN"])
            m_knm = float(r["M_kNm"])
            if abs(n_kn) < 1e-9 and abs(m_knm) < 1e-9:
                continue
            loads.append({
                "load_type": str(r.get("load_type", "ULS")).strip(),
                "N_kN": n_kn,
                "M_kNm": m_knm,
            })
        if not loads:
            print(f"跳过构件 {gid}：无有效荷载（全为 0）")
            continue
        print(f"计算构件 {gid}（{len(loads)} 组荷载）...")
        res.append(run_one(group.iloc[0], loads, main_hw, rtf_dir=rtf_dir))

    if not res:
        raise RuntimeError(f"{excel_path} 未产生任何计算结果")

    result_df = pd.DataFrame(res)
    for col in RESULT_COLUMNS:
        if col not in result_df.columns:
            result_df[col] = None
    result_df = result_df[RESULT_COLUMNS]
    if "distance_m" in result_df.columns:
        result_df = result_df.sort_values(
            ["distance_m", "ID"], kind="stable", na_position="last",
        ).reset_index(drop=True)
    result_df.to_excel(result_path, index=False)
    print(f"  → 配筋汇总: {result_path}（{len(result_df)} 个构件）")
    return result_df


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] in ("--calibrate", "--test-coords"):
        import calibrate_load
        if sys.argv[1] == "--calibrate":
            calibrate_load.run_calibrate()
        else:
            calibrate_load.run_test()
        raise SystemExit(0)

    parser = argparse.ArgumentParser(description="Expert N-M 轴力弯矩验算")
    parser.add_argument(
        "excel", nargs="*",
        help="一个或多个 input_NM.xlsx，按顺序依次计算",
    )
    parser.add_argument(
        "--output-base", default=SCRIPT_DIR,
        help="各 xlsx 输出根目录，实际路径为 {output_base}/{xlsx文件名}/",
    )
    args = parser.parse_args()

    excel_list = [os.path.abspath(p) for p in args.excel] if args.excel else [INPUT_EXCEL]

    print("=" * 50)
    print("Expert 轴力+弯矩验算 (N-M)")
    print("运行前请：打开 Expert → Design 页 → 窗口不要最小化")
    print("运行期间请勿操作鼠标键盘")
    print(f"待计算文件（{len(excel_list)} 个）:")
    for p in excel_list:
        print(f"  - {p}")
    print(f"输出根目录: {os.path.abspath(args.output_base)}")
    if os.path.isfile(LOAD_COORDS_FILE):
        print(f"荷载坐标: 已加载 {LOAD_COORDS_FILE}")
        print(
            f"  行号1 X={LOAD_ROW_INDEX_COL_X}  Type={LOAD_TYPE_COL_X}  "
            f"N={LOAD_N_COL_X}  M={LOAD_M_COL_X}  "
            f"翻页点击行={LOAD_SCROLL_CLICK_ROW + 1}  "
            f"翻页填表Y行={LOAD_OVERFLOW_FILL_ROW + 1}  "
            f"首行Y={LOAD_FIRST_ROW_Y}  行高={LOAD_ROW_HEIGHT}"
        )
    else:
        print("荷载坐标: 未校准！请先运行  cd expert/expert_M && python calibrate_load.py")
        print("          或  python expert_try.py --calibrate")
    print("=" * 50)

    main_hw = get_expert_hwnd()
    main_hw = ensure_expert_visible(main_hw)
    ctrl_map = build_control_map(main_hw)
    print(f"控件映射 OK: b={ctrl_map['edit_b']}, h={ctrl_map['edit_h']}, "
          f"seismic={ctrl_map.get('chk_seismic')}")
    print(f"  荷载表 GXWND={ctrl_map.get('grid_hwnd')}  "
          f"Results框={ctrl_map.get('results_log_hwnd')}（填表时会临时禁用）")

    if len(excel_list) == 1 and not os.path.exists(excel_list[0]):
        pd.DataFrame([{
            "ID": "C01", "b_cm": 25, "h_cm": 50, "d1_cm": 5, "d2_cm": 5,
            "is_column": 0, "load_type": "ULS", "N_kN": 1350, "M_kNm": 88,
        }]).to_excel(excel_list[0], index=False)
        print(f"已生成示例: {excel_list[0]}，请修改后重新运行")
        raise SystemExit(0)

    completed = 0
    for idx, excel_path in enumerate(excel_list, start=1):
        if not os.path.isfile(excel_path):
            print(f"\n[{idx}/{len(excel_list)}] 跳过：未找到 {excel_path}")
            continue

        job_dir, rtf_dir, result_path = resolve_job_paths(excel_path, args.output_base)
        print(f"\n{'=' * 50}")
        print(f"[{idx}/{len(excel_list)}] 输入: {excel_path}")
        print(f"  输出目录: {job_dir}")
        print(f"  RTF → {rtf_dir}/")
        print(f"  汇总 → {result_path}")

        try:
            process_input_excel(excel_path, main_hw, rtf_dir, result_path)
            completed += 1
        except Exception as exc:
            print(f"  ✗ 本文件计算失败，继续下一个: {exc}")
            continue

    print(f"\n{'=' * 50}")
    print(f"全部完成：{completed}/{len(excel_list)} 个文件已计算")
