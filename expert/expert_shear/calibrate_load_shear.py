"""
抗剪版荷载表 + 箍筋输入框坐标校准

用法（在 expert_shear 目录）：
  python calibrate_load_shear.py              # 荷载表 6 点 + 箍筋 4 点
  python calibrate_load_shear.py --stirrup-only   # 仅校准箍筋 4 点（保留已有荷载坐标）
  python calibrate_load_shear.py --test

或：
  python expert_try_shear.py --calibrate
  python expert_try_shear.py --test-coords
"""
import argparse
import json
import os
import time

import win32api
import win32gui

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
COORDS_FILE = os.path.join(SCRIPT_DIR, "load_coords_shear.json")
TARGET_WINDOW = "EXPERT RC - Shear and torsion"

STIRRUP_KEYS = ("N1", "D1", "N2", "D2")
STIRRUP_LABELS = (
    "⑦ Results 区 · n1 根数（第一红框，如 2）",
    "⑧ Results 区 · φ1 直径下拉（第二红框，如 16）",
    "⑨ Results 区 · n2 根数（第三红框，如 10）",
    "⑩ Results 区 · φ2 直径下拉（第四蓝框，如 12）",
)


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


def get_design_hw(main_hw):
    design = []
    win32gui.EnumChildWindows(
        main_hw,
        lambda hw, arr: arr.append(hw) or True
        if win32gui.GetWindowText(hw) == "Design"
        else True,
        design,
    )
    if not design:
        raise RuntimeError("未找到 Design 页，请先切换到 Design")
    return design[0]


def get_gxwnd_rect(main_hw):
    design_hw = get_design_hw(main_hw)
    gxwnds = []
    win32gui.EnumChildWindows(
        design_hw,
        lambda hw, arr: arr.append((win32gui.GetClassName(hw), win32gui.GetWindowRect(hw))) or True
        if win32gui.GetClassName(hw) == "GXWND"
        else True,
        gxwnds,
    )
    if not gxwnds:
        raise RuntimeError("未找到荷载表 GXWND")
    _, rect = max(gxwnds, key=lambda r: (r[1][2] - r[1][0]) * (r[1][3] - r[1][1]))
    return rect


def get_design_rect(main_hw):
    return win32gui.GetWindowRect(get_design_hw(main_hw))


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


def screen_points(gl, gt, dl, dt, coords):
    y1 = gt + coords["LOAD_FIRST_ROW_Y"]
    y2 = y1 + coords["LOAD_ROW_HEIGHT"]
    v_x = coords.get("LOAD_V_COL_X", coords["LOAD_TYPE_COL_X"] + 51)
    n_x = coords.get("LOAD_N_COL_X", v_x + 51)
    exp_x = gl + coords.get("LOAD_EXPAND_BTN_X", n_x + 70)
    exp_y = gt + coords.get("LOAD_EXPAND_BTN_Y", y1 + 80)
    points = [
        ("行号 1（清空用）", gl + coords["LOAD_ROW_INDEX_COL_X"], y1),
        ("Load type 第1行", gl + coords["LOAD_TYPE_COL_X"], y1),
        ("V 第1行", gl + v_x, y1),
        ("N 第1行", gl + n_x, y1),
        ("Load type 第2行（行高）", gl + coords["LOAD_TYPE_COL_X"], y2),
        ("右侧下拉按钮（展开第10行）", exp_x, exp_y),
    ]
    for key in STIRRUP_KEYS:
        xk, yk = f"STIRRUP_{key}_X", f"STIRRUP_{key}_Y"
        if xk in coords and yk in coords:
            points.append((
                f"箍筋 {key}",
                dl + coords[xk],
                dt + coords[yk],
            ))
    return points


def calibrate_load_table(main_hw, coords):
    gl, gt, gr, gb = get_gxwnd_rect(main_hw)
    print("=" * 55)
    print("GXWND 荷载表外框（抗剪 Shear and torsion）")
    print(f"  左上 ({gl}, {gt})   右下 ({gr}, {gb})")
    print(f"  宽={gr - gl}  高={gb - gt}")
    print("=" * 55)
    print("依次校准 6 个点（鼠标移到目标正中心，等倒计时结束）：")

    x_idx, y_idx = pick_point("① 第 1 行 · 左侧行号「1」")
    x_lt, y_lt = pick_point("② 第 1 行 · Load type（Type）白格")
    x_v, _ = pick_point("③ 第 1 行 · V 白格")
    x_n, _ = pick_point("④ 第 1 行 · N 白格")
    x_lt2, y_lt2 = pick_point("⑤ 第 2 行 · Load type 白格（算行高）")
    x_exp, y_exp = pick_point("⑥ 荷载表右侧 · 下拉按钮（展开更多行）")

    coords.update({
        "LOAD_ROW_INDEX_COL_X": x_idx - gl,
        "LOAD_TYPE_COL_X": x_lt - gl,
        "LOAD_V_COL_X": x_v - gl,
        "LOAD_N_COL_X": x_n - gl,
        "LOAD_EXPAND_BTN_X": x_exp - gl,
        "LOAD_EXPAND_BTN_Y": y_exp - gt,
        "LOAD_FIRST_ROW_Y": y_lt - gt,
        "LOAD_ROW_HEIGHT": y_lt2 - y_lt,
    })
    return coords


def calibrate_stirrup(main_hw, coords):
    dl, dt, _, _ = get_design_rect(main_hw)
    print("\n" + "=" * 55)
    print("箍筋 Cadres 输入框（相对 Design 页左上角存偏移）")
    print(f"  Design 左上 ({dl}, {dt})")
    print("=" * 55)
    print("依次校准 4 个点（移到输入框正中心，等倒计时结束）：")
    print("提示：④ φ2 框在 n2=0 时灰色，请先在界面填 n2≠0 再校准第 4 点。")

    picked = {}
    for key, label in zip(STIRRUP_KEYS, STIRRUP_LABELS):
        x, y = pick_point(label)
        picked[f"STIRRUP_{key}_X"] = x - dl
        picked[f"STIRRUP_{key}_Y"] = y - dt
    coords.update(picked)
    return coords


def run_calibrate(stirrup_only=False):
    main_hw = get_expert_hwnd()
    coords = load_coords() or {}

    if not stirrup_only:
        coords = calibrate_load_table(main_hw, coords)
    elif not coords.get("LOAD_FIRST_ROW_Y"):
        print("警告: 尚无荷载表坐标，将仅写入箍筋坐标。")

    coords = calibrate_stirrup(main_hw, coords)
    save_coords(coords)

    print("\n写入的值：")
    for k, v in sorted(coords.items()):
        print(f"  {k} = {v}")
    print("\n运行测试确认：  python calibrate_load_shear.py --test")
    return coords


def run_test():
    main_hw = get_expert_hwnd()
    gl, gt, _, _ = get_gxwnd_rect(main_hw)
    dl, dt, _, _ = get_design_rect(main_hw)
    coords = load_coords()
    if not coords:
        print(f"未找到 {COORDS_FILE}，请先运行: python calibrate_load_shear.py")
        return

    print("=" * 55)
    print("测试模式：鼠标将依次移到各校准点，每点停留 4 秒")
    print("=" * 55)
    for name, x, y in screen_points(gl, gt, dl, dt, coords):
        print(f"  → {name}  ({int(x)}, {int(y)})")
        win32api.SetCursorPos((int(x), int(y)))
        time.sleep(4)
    print("\n测试结束。")


def main():
    parser = argparse.ArgumentParser(description="Expert 抗剪荷载表 + 箍筋坐标校准")
    parser.add_argument("--test", action="store_true", help="测试已保存的坐标")
    parser.add_argument(
        "--stirrup-only",
        action="store_true",
        help="仅校准箍筋 4 点（保留已有荷载表坐标）",
    )
    args = parser.parse_args()
    if args.test:
        run_test()
    else:
        run_calibrate(stirrup_only=args.stirrup_only)
        ans = input("\n是否立即测试坐标？(y/n): ").strip().lower()
        if ans == "y":
            run_test()


if __name__ == "__main__":
    main()
