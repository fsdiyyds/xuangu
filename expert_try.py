import win32gui
import win32con
import win32api
import time
import pandas as pd

# =====================配置区（根据你Expert界面微调，默认适配常规Expert RC2010）=====================
EXPERT_WND_TITLE = "Expert RC 2010"  # Expert窗口标题，SPY++核对
SLEEP_SHORT = 0.35
SLEEP_CALC = 0.8
EXCEL_IN = "input_NM.xlsx"
EXCEL_OUT = "result_配筋汇总.xlsx"
# ==========================================================================================

def get_main_hwnd():
    """获取Expert主窗口句柄"""
    hwnd = win32gui.FindWindow(None, EXPERT_WND_TITLE)
    if hwnd == 0:
        raise Exception("未找到Expert2010窗口，请先手动打开软件！")
    return hwnd

def enum_all_child(hwnd_parent, child_list):
    """递归遍历全部子控件句柄"""
    def callback(hwnd, param):
        child_list.append(hwnd)
        return True
    win32gui.EnumChildWindows(hwnd_parent, callback, None)

def set_edit_text(hwnd_edit, text):
    """编辑框填入数值"""
    win32gui.SendMessage(hwnd_edit, win32con.WM_SETTEXT, 0, str(text))
    time.sleep(SLEEP_SHORT)

def click_button(hwnd_btn):
    """模拟按钮点击"""
    win32gui.PostMessage(hwnd_btn, win32con.WM_LBUTTONDOWN, 0, 0)
    win32gui.PostMessage(hwnd_btn, win32con.WM_LBUTTONUP, 0, 0)
    time.sleep(SLEEP_SHORT)

def get_edit_value(hwnd_edit):
    """读取编辑框结果"""
    buf_len = win32gui.SendMessage(hwnd_edit, win32con.WM_GETTEXTLENGTH, 0, 0) + 1
    buf = win32gui.PyMakeBuffer(buf_len)
    win32gui.SendMessage(hwnd_edit, win32con.WM_GETTEXT, buf_len, buf)
    return str(buf[:-1]).strip()

def auto_calc_one_line(hwnd_main, row):
    """单构件自动配筋：填参数→计算→读As1 As2"""
    child_hwnds = []
    enum_all_child(hwnd_main, child_hwnds)

    # ==========控件索引顺序（Expert固定排布：b、h、d1、d2、column单选、N、M、Calculate、As1、As2）==========
    # 若你软件控件顺序不一致：用SPY++抓控件，修改下面下标
    idx_b = 0
    idx_h = 1
    idx_d1 = 2
    idx_d2 = 3
    idx_radio_col = 4
    idx_N = 5
    idx_M = 6
    idx_btn_calc = 7
    idx_As1 = 8
    idx_As2 = 9

    # 1.填写截面 b h d1 d2(cm)
    set_edit_text(child_hwnds[idx_b], row["b_cm"])
    set_edit_text(child_hwnds[idx_h], row["h_cm"])
    set_edit_text(child_hwnds[idx_d1], row["d1_cm"])
    set_edit_text(child_hwnds[idx_d2], row["d2_cm"])

    # 2.切换梁/柱单选框 is_column=1点column，0默认beam
    if int(row["is_column"]) == 1:
        click_button(child_hwnds[idx_radio_col])

    # 3.填入ULS N(kN)、M(kN·m)
    set_edit_text(child_hwnds[idx_N], row["N_kN"])
    set_edit_text(child_hwnds[idx_M], row["M_kNm"])

    # 4.点击计算Calculate
    click_button(child_hwnds[idx_btn_calc])
    time.sleep(SLEEP_CALC)

    # 5.读取结果 As1(下筋cm²) As2(上筋cm²)
    As1 = get_edit_value(child_hwnds[idx_As1])
    As2 = get_edit_value(child_hwnds[idx_As2])

    return {"ID":row["ID"], "b":row["b_cm"], "h":row["h_cm"],
            "N":row["N_kN"], "M":row["M_kNm"],
            "As1_cm2":As1, "As2_cm2":As2}

if __name__ == "__main__":
    # 读取内力表
    df_in = pd.read_excel(EXCEL_IN)
    res_list = []
    hwnd_expert = get_main_hwnd()

    print(f"共{len(df_in)}个构件开始批量计算...")
    for idx, line in df_in.iterrows():
        print(f"正在计算：{line['ID']}")
        one_res = auto_calc_one_line(hwnd_expert, line)
        res_list.append(one_res)
        # 每个算完清空N/M，避免下一行干扰
        childs = []
        enum_all_child(hwnd_expert, childs)
        set_edit_text(childs[5], "")
        set_edit_text(childs[6], "")
        time.sleep(SLEEP_SHORT)

    # 导出结果
    df_out = pd.DataFrame(res_list)
    df_out.to_excel(EXCEL_OUT, index=False)
    print(f"批量完成！结果已保存→{EXCEL_OUT}")