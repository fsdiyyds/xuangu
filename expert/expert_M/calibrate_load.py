"""
Expert N-M 荷载表坐标校准（expert_M）

用法：
  cd expert/expert_M
  python calibrate_load.py
  python calibrate_load.py --test
  python expert_try.py --calibrate

校准结果保存到 load_coords.json，expert_try.py 会自动读取。
"""
import argparse
import json
import os
import time
import win32api
import win32gui

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
COORDS_FILE = os.path.join(SCRIPT_DIR, "load_coords.json")
TARGET_WINDOW = "EXPERT RC - Combined axial load and bending"


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


def get_gxwnd_rect(main_hw):
    design = []
    win32gui.EnumChildWindows(
        main_hw,
        lambda hw, arr: arr.append(hw) or True
        if win32gui.GetWindowText(hw) == "Design"
        else True,
        design,
    )
    if not design:
        raise RuntimeError("未找到 Design 页")

    gxwnds = []
    win32gui.EnumChildWindows(
        design[0],
        lambda hw, arr: arr.append((win32gui.GetClassName(hw), win32gui.GetWindowRect(hw))) or True
        if win32gui.GetClassName(hw) == "GXWND"
        else True,
        gxwnds,
    )
    if not gxwnds:
        raise RuntimeError("未找到荷载表 GXWND")
    _, rect = max(gxwnds, key=lambda r: (r[1][2] - r[1][0]) * (r[1][3] - r[1][1]))
    return rect


def pick_point(label):
    print(f"\n>>> {label}")
    print("    5 秒内把鼠标移到目标格子正中心，不要点击...")
    for i in range(5, 0, -1):
        print(f"    {i}...", end="\r")
        time.sleep(1)
    x, y = win32api.GetCursorPos()
    print(f"    已记录 ({x}, {y})              ")
    return x, y


def save_coords(coords):
    with open(COORDS_FILE, "w", encoding="utf-8") as f:
        json.dump(coords, f, indent=2, ensure_ascii=False)
    print(f"\n已保存 → {COORDS_FILE}")


def load_coords():
    if not os.path.isfile(COORDS_FILE):
        return None
    with open(COORDS_FILE, encoding="utf-8") as f:
        return json.load(f)


def screen_points(gl, gt, coords):
    y1 = gt + coords["LOAD_FIRST_ROW_Y"]
    y2 = y1 + coords["LOAD_ROW_HEIGHT"]
    n_x = coords.get("LOAD_N_COL_X", coords["LOAD_TYPE_COL_X"] + 102)
    m_x = coords.get("LOAD_M_COL_X", coords["LOAD_TYPE_COL_X"] + 195)
    return [
        ("行号 1（清空用）", gl + coords["LOAD_ROW_INDEX_COL_X"], y1),
        ("Load type 第1行", gl + coords["LOAD_TYPE_COL_X"], y1),
        ("N 第1行", gl + n_x, y1),
        ("M 第1行", gl + m_x, y1),
        ("Load type 第2行（行高）", gl + coords["LOAD_TYPE_COL_X"], y2),
    ]


def run_calibrate():
    main_hw = get_expert_hwnd()
    gl, gt, gr, gb = get_gxwnd_rect(main_hw)

    print("=" * 55)
    print("GXWND 荷载表外框")
    print(f"  左上 ({gl}, {gt})   右下 ({gr}, {gb})")
    print(f"  宽={gr - gl}  高={gb - gt}")
    print("=" * 55)
    print("依次校准 5 个点（鼠标移到格子正中心，等倒计时结束）：")

    x_idx, y_idx = pick_point("① 第 1 行 · 左侧行号「1」")
    x_lt, y_lt = pick_point("② 第 1 行 · Load type 白格")
    x_n, _ = pick_point("③ 第 1 行 · N 白格")
    x_m, _ = pick_point("④ 第 1 行 · M 白格")
    x_lt2, y_lt2 = pick_point("⑤ 第 2 行 · Load type 白格（算行高）")

    coords = {
        "LOAD_ROW_INDEX_COL_X": x_idx - gl,
        "LOAD_TYPE_COL_X": x_lt - gl,
        "LOAD_N_COL_X": x_n - gl,
        "LOAD_M_COL_X": x_m - gl,
        "LOAD_FIRST_ROW_Y": y_lt - gt,
        "LOAD_ROW_HEIGHT": y_lt2 - y_lt,
    }
    save_coords(coords)

    print("\n写入的值：")
    for k, v in coords.items():
        print(f"  {k} = {v}")
    print("\n运行测试确认：  python calibrate_load.py --test")
    return coords, (gl, gt)


def run_test():
    main_hw = get_expert_hwnd()
    gl, gt, _, _ = get_gxwnd_rect(main_hw)
    coords = load_coords()
    if not coords:
        print(f"未找到 {COORDS_FILE}，请先运行: python calibrate_load.py")
        return

    print("=" * 55)
    print("测试模式：鼠标将依次移到各校准点，每点停留 4 秒")
    print("请观察是否对准格子中心（不对请重新 calibrate）")
    print("=" * 55)
    for name, x, y in screen_points(gl, gt, coords):
        print(f"  → {name}  ({int(x)}, {int(y)})")
        win32api.SetCursorPos((int(x), int(y)))
        time.sleep(4)
    print("\n测试结束。")


def main():
    parser = argparse.ArgumentParser(description="Expert 荷载表坐标校准")
    parser.add_argument("--test", action="store_true", help="测试已保存的坐标")
    args = parser.parse_args()
    if args.test:
        run_test()
    else:
        coords, _ = run_calibrate()
        ans = input("\n是否立即测试坐标？(y/n): ").strip().lower()
        if ans == "y":
            run_test()


if __name__ == "__main__":
    main()
