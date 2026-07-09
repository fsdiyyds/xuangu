"""
Expert RC 2010 控件诊断工具
用法：手动打开 Expert → 进入 Combined axial load and bending → Design 页 → 运行本脚本
输出 control_map.txt，用于核对控件类名与屏幕坐标顺序
"""
import win32gui
import win32con

TARGET_WINDOW = "EXPERT RC - Combined axial load and bending"


def get_expert_hwnd():
    hwnds = []

    def cb(hw, arr):
        if win32gui.IsWindowVisible(hw) and TARGET_WINDOW in win32gui.GetWindowText(hw):
            arr.append(hw)
        return True

    win32gui.EnumWindows(cb, hwnds)
    if not hwnds:
        raise RuntimeError(f"未找到窗口：{TARGET_WINDOW}")
    return hwnds[0]


def get_design_page_hwnd(main_hwnd):
    hwnds = []

    def cb(hw, arr):
        if win32gui.GetWindowText(hw) == "Design":
            arr.append(hw)
        return True

    win32gui.EnumChildWindows(main_hwnd, cb, hwnds)
    if not hwnds:
        raise RuntimeError("未找到 Design 子窗口，请确认已切换到 Design 页")
    return hwnds[0]


def collect_controls(parent):
    rows = []

    def cb(hw, _):
        cls = win32gui.GetClassName(hw)
        text = win32gui.GetWindowText(hw)
        ctrl_id = win32gui.GetDlgCtrlID(hw)
        left, top, right, bottom = win32gui.GetWindowRect(hw)
        rows.append({
            "hwnd": hw,
            "class": cls,
            "text": text,
            "id": ctrl_id,
            "left": left,
            "top": top,
            "right": right,
            "bottom": bottom,
        })
        return True

    win32gui.EnumChildWindows(parent, cb, None)
    rows.sort(key=lambda r: (r["top"], r["left"]))
    return rows


def enum_descendants(parent, depth=0):
    rows = []

    def cb(hw, _):
        cls = win32gui.GetClassName(hw)
        text = win32gui.GetWindowText(hw)
        ctrl_id = win32gui.GetDlgCtrlID(hw)
        left, top, right, bottom = win32gui.GetWindowRect(hw)
        rows.append({
            "depth": depth,
            "hwnd": hw,
            "class": cls,
            "text": text,
            "id": ctrl_id,
            "left": left,
            "top": top,
            "right": right,
            "bottom": bottom,
        })
        rows.extend(enum_descendants(hw, depth + 1))
        return True

    win32gui.EnumChildWindows(parent, cb, None)
    return rows


def format_rows(title, rows):
    lines = [
        title,
        f"{'idx':>4}  {'id':>6}  {'class':<24}  {'text':<30}  rect(left,top,right,bottom)",
        "-" * 110,
    ]
    for i, r in enumerate(rows):
        text = r["text"].replace("\n", " ")[:30]
        rect = f"({r['left']},{r['top']},{r['right']},{r['bottom']})"
        indent = "  " * r.get("depth", 0)
        lines.append(
            f"{i:>4}  {r['id']:>6}  {r['class']:<24}  {indent}{text:<30}  {rect}"
        )
    return lines


def main():
    main_hw = get_expert_hwnd()
    design_hw = get_design_page_hwnd(main_hw)
    rows = collect_controls(design_hw)

    lines = [
        f"主窗口: {win32gui.GetWindowText(main_hw)} (hwnd={main_hw})",
        f"Design页: hwnd={design_hw}",
        f"子控件数量: {len(rows)}",
        "",
    ]
    lines.extend(format_rows("【Design 直接子控件】", rows))

    gxwnd_rows = [r for r in rows if r["class"] == "GXWND"]
    if gxwnd_rows:
        gxwnd_hw = gxwnd_rows[0]["hwnd"]
        inner = enum_descendants(gxwnd_hw)
        inner.sort(key=lambda r: (r["top"], r["left"]))
        lines.extend(["", f"【GXWND 荷载表内部控件】(hwnd={gxwnd_hw})", ""])
        lines.extend(format_rows("", inner))

    output = "\n".join(lines)
    print(output)

    out_file = "control_map.txt"
    with open(out_file, "w", encoding="utf-8") as f:
        f.write(output)
    print(f"\n已保存到 {out_file}")


if __name__ == "__main__":
    main()
