import win32gui
import win32con
import win32api
import time
import pandas as pd
import re
import shutil
import os

# ===================== 配置 =====================
SLEEP_SHORT = 0.6
SLEEP_CALC = 1.5
SLEEP_FILE = 1.0
EXCEL_IN = "input_NM.xlsx"
EXCEL_OUT = "result_配筋.xlsx"
RTF_SOURCE = "rc_note.rtf"
# =================================================

def get_expert_hwnd():
    target = "EXPERT RC - Combined axial load and bending"
    lst = []
    def cb(hw, arr):
        if win32gui.IsWindowVisible(hw):
            t = win32gui.GetWindowText(hw)
            if target in t:
                arr.append(hw)
    win32gui.EnumWindows(cb, lst)
    return lst[0]

def all_children(hw):
    c = []
    def ec(h, p):
        c.append(h)
        return True
    win32gui.EnumChildWindows(hw, ec, None)
    return c

def set_val(hw, txt):
    win32gui.SendMessage(hw, win32con.WM_SETTEXT, 0, str(txt))
    time.sleep(SLEEP_SHORT)

def click(hw):
    win32gui.PostMessage(hw, win32con.WM_LBUTTONDOWN, 0,0)
    win32gui.PostMessage(hw, win32con.WM_LBUTTONUP, 0,0)
    time.sleep(SLEEP_SHORT)

def choose_uls(cbo):
    click(cbo)
    time.sleep(0.3)
    win32api.keybd_event(win32con.VK_DOWN,0,0,0)
    win32api.keybd_event(win32con.VK_RETURN,0,0,0)
    time.sleep(SLEEP_SHORT)

def read_rtf_and_save_as(new_name):
    if os.path.exists(RTF_SOURCE):
        dest = f"Note_{new_name}.rtf"
        shutil.copy2(RTF_SOURCE, dest)
        time.sleep(SLEEP_FILE)
        return dest
    return "无文件"

def calc_one(main_hw, row):
    ch = all_children(main_hw)

    # ===================== 【终极控件索引】 =====================
    cbo_load = ch[3]      # Load Type 下拉
    ed_N     = ch[6]      # ULS 行 N 输入框
    ed_M     = ch[7]      # ULS 行 M 输入框
    rad_beam = ch[18]
    rad_col  = ch[19]
    chk_seis = ch[20]     # Seismic detailing
    ed_b     = ch[22]
    ed_h     = ch[23]
    ed_d1    = ch[24]
    ed_d2    = ch[25]
    btn_calc = ch[26]
    btn_note = ch[-1]
    # ============================================================

    # 1 选 ULS
    choose_uls(cbo_load)

    # 2 勾选 Seismic
    click(chk_seis)

    # 3 截面
    set_val(ed_b, row["b_cm"])
    set_val(ed_h, row["h_cm"])
    set_val(ed_d1, row["d1_cm"])
    set_val(ed_d2, row["d2_cm"])

    # 4 梁柱
    if int(row.get("is_column", 0)) == 1:
        click(rad_col)
    else:
        click(rad_beam)

    # 5 填入 N / M —— 【这次一定能看见！】
    set_val(ed_N, row["N_kN"])
    set_val(ed_M, row["M_kNm"])

    # 6 计算
    click(btn_calc)
    time.sleep(SLEEP_CALC)

    # 7 点 Note → 生成 rc_note.rtf
    click(btn_note)
    time.sleep(SLEEP_FILE)

    # 8 另存为
    rtf_path = read_rtf_and_save_as(row["ID"])

    # 9 清空
    set_val(ed_N, "")
    set_val(ed_M, "")

    return {
        "ID": row["ID"],
        "b_cm": row["b_cm"],
        "h_cm": row["h_cm"],
        "is_column": row.get("is_column",0),
        "N_kN": row["N_kN"],
        "M_kNm": row["M_kNm"],
        "Note文件": rtf_path
    }

# ===================== 主程序 =====================
if __name__ == "__main__":
    df = pd.read_excel(EXCEL_IN)
    hw = get_expert_hwnd()
    res = []
    for _, row in df.iterrows():
        print(f"计算：{row['ID']}")
        res.append(calc_one(hw, row))
    pd.DataFrame(res).to_excel(EXCEL_OUT, index=False)
    print("✅ 全部完成！每个构件都生成独立 Note 文件")