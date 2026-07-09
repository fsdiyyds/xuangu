"""探测 rcba 弹框结构。用法：先让 Expert 弹出 rcba 对话框，再运行本脚本。"""
import win32gui
import win32con
import sys

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


def walk(hwnd, depth=0, lines=None):
    if lines is None:
        lines = []
    try:
        cls = win32gui.GetClassName(hwnd)
        text = win32gui.GetWindowText(hwnd)
        vis = win32gui.IsWindowVisible(hwnd)
        enabled = win32gui.EnableWindow(hwnd, True)  # no-op read
        rect = win32gui.GetWindowRect(hwnd)
        ctrl_id = win32gui.GetWindowLong(hwnd, win32con.GWL_ID)
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
        lines.append(
            f"{'  ' * depth}[{hwnd}] class={cls!r} id={ctrl_id} "
            f"vis={vis} text={text!r} rect={rect} style=0x{style:X}"
        )

        def on_child(ch, _):
            walk(ch, depth + 1, lines)
            return True

        win32gui.EnumChildWindows(hwnd, on_child, None)
    except Exception as e:
        lines.append(f"{'  ' * depth}ERROR hwnd={hwnd}: {e}")
    return lines


def find_rcba_windows():
    found = []
    def enum_top(hw, _):
        if win32gui.IsWindowVisible(hw):
            title = win32gui.GetWindowText(hw)
            cls = win32gui.GetClassName(hw)
            if title == "rcba" or "rcba" in title.lower() or cls == "#32770":
                t = win32gui.GetWindowText(hw)
                if t == "rcba" or t == "":
                    found.append(hw)
        return True
    win32gui.EnumWindows(enum_top, None)
    # 也按标题精确匹配
    exact = []
    win32gui.EnumWindows(
        lambda hw, _: exact.append(hw) or True
        if win32gui.IsWindowVisible(hw) and win32gui.GetWindowText(hw) == "rcba"
        else True,
        None,
    )
    return list(dict.fromkeys(exact + found))


def dump_buttons(hwnd):
    buttons = []
    def on_btn(hw, _):
        cls = win32gui.GetClassName(hw)
        if cls in ("Button", "TButton"):
            text = win32gui.GetWindowText(hw)
            cid = win32gui.GetWindowLong(hw, win32con.GWL_ID)
            rect = win32gui.GetWindowRect(hw)
            buttons.append((hw, cls, text, cid, rect))
        return True
    win32gui.EnumChildWindows(hwnd, on_btn, None)
    return buttons


def try_click_yes(dlg):
    """与 expert_try_biaxial.dismiss_rcba_dialog 相同逻辑"""
    from expert_try_biaxial import dismiss_rcba_dialog

    print("\n尝试 dismiss_rcba_dialog() ...")
    ok = dismiss_rcba_dialog()
    still = find_rcba_windows()
    print(f"  结果: dismissed={ok}, 仍可见 rcba 窗口={still}")
    return ok and not still


def main():
    click_yes = "--click-yes" in sys.argv
    print("=" * 60)
    print("rcba 对话框探测" + (" + 点「是」" if click_yes else ""))
    print("=" * 60)

    rcba_list = find_rcba_windows()
    if not rcba_list:
        print("未找到标题为 rcba 的可见窗口。")
        print("尝试列出所有可见顶层窗口标题（含 rcba / #32770）：")
        tops = []
        win32gui.EnumWindows(
            lambda hw, _: tops.append(
                (win32gui.GetWindowText(hw), win32gui.GetClassName(hw), hw)
            ) or True
            if win32gui.IsWindowVisible(hw)
            else True,
            None,
        )
        for title, cls, hw in tops:
            if not title and cls != "#32770":
                continue
            if "rcba" in (title or "").lower() or cls == "#32770":
                print(f"  title={title!r} class={cls!r} hwnd={hw}")
        sys.exit(1)

    for i, dlg in enumerate(rcba_list):
        title = win32gui.GetWindowText(dlg)
        cls = win32gui.GetClassName(dlg)
        print(f"\n--- 窗口 #{i + 1} hwnd={dlg} title={title!r} class={cls!r} ---")
        print(f"窗口矩形: {win32gui.GetWindowRect(dlg)}")
        print("\n完整控件树:")
        for line in walk(dlg):
            print(line)

        print("\n按钮列表:")
        for hw, bcls, text, cid, rect in dump_buttons(dlg):
            cx = (rect[0] + rect[2]) // 2
            cy = (rect[1] + rect[3]) // 2
            id_hint = ""
            if cid == win32con.IDYES:
                id_hint = " (=IDYES，应用 WM_COMMAND 点「是」)"
            elif cid == win32con.IDNO:
                id_hint = " (=IDNO)"
            print(
                f"  hwnd={hw} class={bcls!r} id={cid} text={text!r} "
                f"rect={rect} center=({cx},{cy}){id_hint}"
            )
        yes_btn = win32gui.GetDlgItem(dlg, win32con.IDYES)
        print(f"\nGetDlgItem(dlg, IDYES={win32con.IDYES}) -> {yes_btn}")

        # 探测父窗口链
        print("\n父窗口链:")
        p = dlg
        for d in range(8):
            p = win32gui.GetParent(p)
            if not p:
                break
            print(
                f"  depth={d} hwnd={p} "
                f"title={win32gui.GetWindowText(p)!r} "
                f"class={win32gui.GetClassName(p)!r}"
            )

    if click_yes and rcba_list:
        try_click_yes(rcba_list[0])

    print("\n" + "=" * 60)
    print("完成。")
    if not click_yes:
        print("若需测试自动点「是」: python inspect_rcba_dialog.py --click-yes")


if __name__ == "__main__":
    main()
