import os
import time

import pyautogui
import pyperclip
import xlrd

# 减少 pyautogui 默认动作间隔
pyautogui.PAUSE = 0.05
pyautogui.FAILSAFE = True

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ================= 全局缓存 =================
coord_box_pos = None   # 上方坐标多行输入框
name_box_pos = None    # 下方名称单行输入框
section_cache = []     # 预解析的断面坐标，避免每次读文件
section_idx = 0
name_count = 0

IMG_COORD = os.path.join(SCRIPT_DIR, "S1.png")
IMG_NAME = os.path.join(SCRIPT_DIR, "S2.png")
WORDS_FILE = os.path.join(SCRIPT_DIR, "words.txt")
CMD_FILE = os.path.join(SCRIPT_DIR, "cmd.xls")

LOCATE_TIMEOUT = 60    # 定位输入框最长等待秒数
PASTE_WAIT = 0.3
# Midas 点选节点后常见显示： [ 36.773810, -31.631600, 0.000000 ]
# 小数位需与模型节点坐标一致，四舍五入到 4 位会导致“坐标号”校验失败
COORD_DECIMALS = 6


def locate_box(img_path, label):
    """定位输入框，超时后给出明确提示"""
    deadline = time.time() + LOCATE_TIMEOUT
    while time.time() < deadline:
        pos = pyautogui.locateCenterOnScreen(img_path, confidence=0.9)
        if pos:
            print(f"已缓存{label} X:{pos.x}, Y:{pos.y}")
            return pos
        time.sleep(0.2)
    raise RuntimeError(
        f"未找到{label}，请确认 Midas 对话框已打开，且截图 {os.path.basename(img_path)} 与当前界面一致"
    )


def init_two_box_pos():
    global coord_box_pos, name_box_pos
    if coord_box_pos is None:
        coord_box_pos = locate_box(IMG_COORD, "坐标框")
    if name_box_pos is None:
        name_box_pos = locate_box(IMG_NAME, "名称框")


def clear_by_pos(pos):
    """Midas 坐标框多为自定义控件，Ctrl+A 常无效，用全选+删除多种方式确保清空"""
    pyautogui.click(pos.x, pos.y, duration=0.05)
    time.sleep(0.15)
    pyautogui.hotkey("ctrl", "home")
    time.sleep(0.05)
    pyautogui.hotkey("ctrl", "shift", "end")
    time.sleep(0.05)
    pyautogui.press("delete")
    time.sleep(0.05)
    # 兜底：再删一遍末尾残留，避免旧坐标与新坐标混在一起
    for _ in range(8):
        pyautogui.press("delete")
        time.sleep(0.01)
    time.sleep(0.1)


def clear_all_box():
    clear_by_pos(coord_box_pos)
    clear_by_pos(name_box_pos)


def paste_text(text):
    pyperclip.copy(text)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(PASTE_WAIT)


def mouseClick(clickTimes, lOrR, img, reTry):
    img_path = img if os.path.isabs(img) else os.path.join(SCRIPT_DIR, img)
    if reTry == 1:
        while True:
            location = pyautogui.locateCenterOnScreen(img_path, confidence=0.9)
            if location is not None:
                pyautogui.click(
                    location.x, location.y,
                    clicks=clickTimes, interval=0.15, duration=0.1, button=lOrR,
                )
                break
            print(f"未找到匹配图片 {img}，0.1 秒后重试")
            time.sleep(0.1)
    elif reTry == -1:
        while True:
            location = pyautogui.locateCenterOnScreen(img_path, confidence=0.9)
            if location is not None:
                pyautogui.click(
                    location.x, location.y,
                    clicks=clickTimes, interval=0.15, duration=0.1, button=lOrR,
                )
            time.sleep(0.1)
    elif reTry > 1:
        for _ in range(reTry):
            location = pyautogui.locateCenterOnScreen(img_path, confidence=0.9)
            if location is not None:
                pyautogui.click(
                    location.x, location.y,
                    clicks=clickTimes, interval=0.15, duration=0.1, button=lOrR,
                )
            time.sleep(0.1)


def dataCheck(sheet1):
    checkCmd = True
    if sheet1.nrows < 2:
        print("没数据啊哥")
        return False
    i = 1
    while i < sheet1.nrows:
        cmdType = sheet1.row(i)[0]
        if cmdType.ctype != 2 or cmdType.value not in (1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0):
            print("第", i + 1, "行,第1列数据有毛病")
            checkCmd = False
        cmdValue = sheet1.row(i)[1]
        if cmdType.value in (1.0, 2.0, 3.0):
            if cmdValue.ctype != 1:
                print("第", i + 1, "行,第2列数据有毛病")
                checkCmd = False
        if cmdType.value == 4.0 and cmdValue.ctype == 0:
            print("第", i + 1, "行,第2列数据有毛病")
            checkCmd = False
        if cmdType.value in (5.0, 6.0) and cmdValue.ctype != 2:
            print("第", i + 1, "行,第2列数据有毛病")
            checkCmd = False
        i += 1
    return checkCmd


def format_midas_coord(x, y, z):
    """生成与 Midas 点选节点后一致的 GCS 坐标行"""
    return f"[ {x:.{COORD_DECIMALS}f}, {y:.{COORD_DECIMALS}f}, {z:.{COORD_DECIMALS}f} ]"


def load_sections():
    """启动时一次性解析 words.txt，循环时直接取用"""
    if not os.path.isfile(WORDS_FILE):
        raise FileNotFoundError(f"找不到坐标文件: {WORDS_FILE}")

    with open(WORDS_FILE, "r", encoding="utf-8") as f:
        all_lines = [line.replace("|", "").strip() for line in f if line.strip()]

    sections = []
    temp = []
    for line in all_lines:
        if "号断面" in line:
            if temp:
                sections.append(temp)
                temp = []
        else:
            temp.append(line)
    if temp:
        sections.append(temp)

    if not sections:
        raise ValueError(f"{WORDS_FILE} 中未解析到任何断面坐标")

    parsed = []
    for sec_lines in sections:
        coords = []
        for line in sec_lines:
            try:
                xs, ys, zs = line.split(",")
                x = float(xs.strip())
                y = float(ys.strip())
                z = float(zs.strip())
                coords.append(format_midas_coord(x, y, z))
            except ValueError:
                continue
        if coords:
            parsed.append("\r\n".join(coords))

    print(f"已加载 {len(parsed)} 组断面坐标")
    if parsed:
        print("首组坐标格式示例（请与 Midas 中手动点选一条对比）:")
        print(parsed[0])
    return parsed


def next_section_text():
    global section_idx
    text = section_cache[section_idx]
    section_idx = (section_idx + 1) % len(section_cache)
    return text


def next_name():
    global name_count
    name_count += 1
    return f"{name_count:02d}"


def get_retry(row):
    if row[2].ctype == 2 and row[2].value != 0:
        return int(row[2].value)
    return 1


def mainWork(sheet1):
    i = 1
    while i < sheet1.nrows:
        row = sheet1.row(i)
        cmdType = row[0]
        if cmdType.value in (1.0, 2.0, 3.0):
            img = row[1].value
            reTry = get_retry(row)
            clicks = 2 if cmdType.value == 2.0 else 1
            button = "right" if cmdType.value == 3.0 else "left"
            mouseClick(clicks, button, img, reTry)
            action = {1.0: "单击左键", 2.0: "双击左键", 3.0: "右键"}[cmdType.value]
            print(action, img)
        elif cmdType.value == 4.0:
            clear_by_pos(coord_box_pos)
            inputValue = next_section_text()
            print(inputValue)
            paste_text(inputValue)
            print("输入:", inputValue)
        elif cmdType.value == 5.0:
            waitTime = row[1].value
            time.sleep(waitTime)
            print("等待", waitTime, "秒")
        elif cmdType.value == 6.0:
            scroll = int(row[1].value)
            pyautogui.scroll(scroll)
            print("滚轮滑动", scroll, "距离")
        elif cmdType.value == 7.0:
            clear_by_pos(name_box_pos)
            inputValue = next_name()
            print(inputValue)
            paste_text(inputValue)
            print("输入:", inputValue)
        i += 1


if __name__ == "__main__":
    print("欢迎使用不高兴就喝水牌 RPA~")
    print("用途：Midas 局部方向内力合力 — 按坐标批量添加截面")

    if not os.path.isfile(CMD_FILE):
        print(f"找不到指令文件: {CMD_FILE}")
        raise SystemExit(1)

    wb = xlrd.open_workbook(filename=CMD_FILE)
    sheet1 = wb.sheet_by_index(0)

    if not dataCheck(sheet1):
        print("输入有误或者已经退出!")
        raise SystemExit(1)

    section_cache.extend(load_sections())

    key = input("选择功能: 1.做一次  2.循环到死\n")
    init_two_box_pos()

    if key == "1":
        mainWork(sheet1)
        clear_all_box()
    elif key == "2":
        while True:
            time.sleep(0.5)
            mainWork(sheet1)
            clear_all_box()
    else:
        print("无效选项，已退出")
