"""
AutoCAD 2018 — 二维多段线顶点引线标注

功能：
  1. 读取 Excel 第一列 ID（如 S1、T1）
  2. 连接已打开的 AutoCAD
  3. 按 CBDI 参数（比例 0.015、字高 10、箭头 10 等）在每个顶点手工绘制：
     空心箭头（尖端指向多段线顶点）+ 引线多段线 + 文字
  4. 标注名 {ID}-01 … {ID}-nn；坐标写入 output.xlsx

用法：
  cd autocad_mleader
  python annotate_polylines.py --drawing 测试1
  python annotate_polylines.py --pick-label   # 点选箭头多段线+文字，保存样式
  # 存在 label_style.json 时自动加载上次样式

依赖：pip install -r requirements.txt
"""
import argparse
import ctypes
import json
import math
import os
import sys
import traceback
import time
import winreg
from typing import Optional

import pandas as pd
import pythoncom
import win32gui
from win32com.client import Dispatch, VARIANT
import win32com.client

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_EXCEL = os.path.join(SCRIPT_DIR, "input_ids.xlsx")
OUTPUT_EXCEL = os.path.join(SCRIPT_DIR, "output.xlsx")
LABEL_STYLE_JSON = os.path.join(SCRIPT_DIR, "label_style.json")

# 点选 / 解析箭头多段线时：第 3 个点（索引 2）为箭头尖端
ARROW_TIP_VERTEX_INDEX = 2
LABEL_POLYLINE_NAMES = frozenset({"AcDbPolyline"})
LABEL_TEXT_NAMES = frozenset({"AcDbMText", "AcDbText"})

# 支持选取的多段线类型
POLYLINE_OBJECT_NAMES = frozenset({
    "AcDbPolyline",       # 轻量多段线 LWPolyline
    "AcDb2dPolyline",     # 二维多段线
})

# 用于从图中读取参考标注参数（可选）
MLEADER_OBJECT_NAME = "AcDbMLeader"

# 默认标注参数（名义尺寸；图面实际 = 名义 × scale）
DEFAULT_STYLE_INFO = {
    "style": "- CBDI Standard",
    "scale": 0.015,
    "text_height": 10.0,
    "arrowhead_size": 10.0,
    "landing_distance": 1.44,
    "text_landing_gap": 8.0,
    "text_style": "- CBDI 2.5mm",
}

# 刚性标注整体：尖端=第3点；u 沿 P6→P7（与文字平行）
ARROW_SIDE_LEN = 0.152
TRACED_SEG67 = (0.1541, 0.5087)
# 描摹参考点，用于保持 P1→P6 沿 u 的投影长度
TRACED_REF_P1 = (-0.1259, 0.0820)
TRACED_REF_P6 = (-0.9403, 0.4402)
TEXT_OFFSET_FROM_TIP = (-0.7015, 0.8777)
TEXT_ROTATION_CANONICAL = math.radians(261.0)
AVOIDANCE_ROT_STEP_DEG = 30

_CANONICAL_RIGID = None


def _unit(dx: float, dy: float) -> tuple:
    length = math.hypot(dx, dy) or 1.0
    return dx / length, dy / length


def _rot2d(x: float, y: float, cos_r: float, sin_r: float) -> tuple:
    return x * cos_r - y * sin_r, x * sin_r + y * cos_r


def build_canonical_rigid_model() -> dict:
    """
    构建相对尖端(原点)的刚性整体：
    - 前5点：等边三角形轮廓 P1→P2→尖端→P4→P1(P5)
    - P5→P6 沿 u，与底边 P1-P4 垂直
    - P6→P7 沿 u，与文字平行
    """
    ux, uy = _unit(*TRACED_SEG67)
    s = ARROW_SIDE_LEN
    ang_b = math.atan2(-uy, -ux)
    # 等边三角形三顶点：尖端 + 两翼（夹角 60°）
    w1 = (
        s * math.cos(ang_b + 2 * math.pi / 3),
        s * math.sin(ang_b + 2 * math.pi / 3),
    )
    w2 = (
        s * math.cos(ang_b + math.pi / 3),
        s * math.sin(ang_b + math.pi / 3),
    )
    p1 = w1
    p2 = (w1[0] * 0.93, w1[1] * 0.93)
    tip = (0.0, 0.0)
    p4 = w2

    leg_u = (
        (TRACED_REF_P6[0] - TRACED_REF_P1[0]) * ux
        + (TRACED_REF_P6[1] - TRACED_REF_P1[1]) * uy
    )
    p6 = (p1[0] + ux * leg_u, p1[1] + uy * leg_u)
    p7 = (p6[0] + TRACED_SEG67[0], p6[1] + TRACED_SEG67[1])

    poly_offsets = [p1, p2, tip, p4, p1, p6, p7]
    return {
        "poly_offsets": poly_offsets,
        "text_offset": TEXT_OFFSET_FROM_TIP,
        "text_rot": TEXT_ROTATION_CANONICAL,
        "u_can": (ux, uy),
        "ang_can": math.atan2(TRACED_SEG67[1], TRACED_SEG67[0]),
    }


def get_canonical_rigid() -> dict:
    global _CANONICAL_RIGID
    if _CANONICAL_RIGID is None:
        _CANONICAL_RIGID = build_canonical_rigid_model()
    return _CANONICAL_RIGID


def set_rigid_label_model(model: dict) -> None:
    global _CANONICAL_RIGID
    _CANONICAL_RIGID = model


def get_lwpolyline_points(entity) -> list:
    """轻量多段线顶点 [(x,y), ...]。"""
    coords = list(entity.Coordinates)
    pts = []
    for i in range(0, len(coords), 2):
        pts.append((float(coords[i]), float(coords[i + 1])))
    return pts


def read_text_entity_placement(text_ent) -> dict:
    """读取 MText / Text 插入点、旋转、字高、样式。"""
    name = text_ent.ObjectName
    info = {"object_name": name}
    p = text_ent.InsertionPoint
    info["x"] = float(p[0])
    info["y"] = float(p[1])
    info["rotation"] = float(text_ent.Rotation)
    info["height"] = float(text_ent.Height)
    try:
        info["style"] = str(text_ent.StyleName)
    except Exception:
        info["style"] = None
    if name == "AcDbMText":
        try:
            info["attachment"] = int(text_ent.AttachmentPoint)
        except Exception:
            info["attachment"] = None
    return info


def build_rigid_model_from_entities(poly_ent, text_ent) -> dict:
    """
    从图中箭头多段线 + 文字构建刚性标注模型（相对尖端坐标）。
    多段线第 3 个点为箭头尖端。
    """
    pts = get_lwpolyline_points(poly_ent)
    if len(pts) < 4:
        raise ValueError(f"箭头多段线至少需要 4 个点，当前 {len(pts)} 个")
    if ARROW_TIP_VERTEX_INDEX >= len(pts):
        raise ValueError(
            f"多段线点数不足：第 3 点（索引 {ARROW_TIP_VERTEX_INDEX}）应为箭头尖端"
        )

    tip_x, tip_y = pts[ARROW_TIP_VERTEX_INDEX]
    poly_offsets = [(x - tip_x, y - tip_y) for x, y in pts]

    txt = read_text_entity_placement(text_ent)
    text_offset = (txt["x"] - tip_x, txt["y"] - tip_y)

    p6 = poly_offsets[-2]
    p7 = poly_offsets[-1]
    seg_x = p7[0] - p6[0]
    seg_y = p7[1] - p6[1]
    if math.hypot(seg_x, seg_y) < 1e-9:
        seg_x, seg_y = math.cos(txt["rotation"]), math.sin(txt["rotation"])
    ux, uy = _unit(seg_x, seg_y)

    return {
        "poly_offsets": poly_offsets,
        "text_offset": text_offset,
        "text_rot": txt["rotation"],
        "text_height": txt["height"],
        "text_style": txt.get("style"),
        "text_attachment": txt.get("attachment"),
        "u_can": (ux, uy),
        "ang_can": math.atan2(uy, ux),
        "tip_index": ARROW_TIP_VERTEX_INDEX,
        "source": "picked",
    }


def rigid_model_to_json(model: dict) -> dict:
    return {
        "poly_offsets": [list(p) for p in model["poly_offsets"]],
        "text_offset": list(model["text_offset"]),
        "text_rot": model["text_rot"],
        "text_height": model.get("text_height"),
        "text_style": model.get("text_style"),
        "text_attachment": model.get("text_attachment"),
        "tip_index": model.get("tip_index", ARROW_TIP_VERTEX_INDEX),
        "source": model.get("source", "picked"),
    }


def load_rigid_model_from_json(path: str) -> dict:
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    poly_offsets = [tuple(p) for p in data["poly_offsets"]]
    text_offset = tuple(data["text_offset"])
    tip_idx = int(data.get("tip_index", ARROW_TIP_VERTEX_INDEX))
    if tip_idx >= len(poly_offsets):
        raise ValueError(f"样式文件 tip_index={tip_idx} 超出点数 {len(poly_offsets)}")

    p6 = poly_offsets[-2]
    p7 = poly_offsets[-1]
    seg_x = p7[0] - p6[0]
    seg_y = p7[1] - p6[1]
    ux, uy = _unit(seg_x, seg_y)

    return {
        "poly_offsets": poly_offsets,
        "text_offset": text_offset,
        "text_rot": float(data["text_rot"]),
        "text_height": data.get("text_height"),
        "text_style": data.get("text_style"),
        "text_attachment": data.get("text_attachment"),
        "u_can": (ux, uy),
        "ang_can": math.atan2(uy, ux),
        "tip_index": tip_idx,
        "source": data.get("source", "file"),
    }


def save_rigid_label_style(model: dict, path: str = LABEL_STYLE_JSON) -> str:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(rigid_model_to_json(model), f, indent=2, ensure_ascii=False)
    return path


def merge_style_from_rigid_model(style_info: dict, model: dict) -> dict:
    """将点选/文件中的文字字高、样式并入 style_info。"""
    merged = dict(style_info)
    if model.get("text_height") is not None:
        merged["drawing_text_height"] = float(model["text_height"])
    if model.get("text_style"):
        merged["text_style"] = model["text_style"]
    return merged


def pick_and_save_label_style(acad) -> dict:
    """点选箭头多段线 + 文字，记录刚性标注样式并保存。"""
    print("  → 请点选【箭头引线多段线】（第 3 个点为箭头尖端）")
    poly, _ = pick_entity(
        acad, "点选箭头引线多段线", LABEL_POLYLINE_NAMES, use_filter=False,
    )
    print("  → 请点选对应的【标注文字】")
    text_ent, _ = pick_entity(
        acad, "点选标注文字", LABEL_TEXT_NAMES, use_filter=False,
    )
    model = build_rigid_model_from_entities(poly, text_ent)
    path = save_rigid_label_style(model)
    set_rigid_label_model(model)
    print(
        f"  → 已记录标注样式：{len(model['poly_offsets'])} 个多段线点，"
        f"尖端=第 {ARROW_TIP_VERTEX_INDEX + 1} 点，已保存 {path}"
    )
    return model


def resolve_rigid_label_style(acad, pick_label: bool = False) -> Optional[dict]:
    """
    加载刚性标注几何：--pick-label 手选 → label_style.json → 内置默认。
    """
    if pick_label:
        return pick_and_save_label_style(acad)

    if os.path.isfile(LABEL_STYLE_JSON):
        try:
            model = load_rigid_model_from_json(LABEL_STYLE_JSON)
            set_rigid_label_model(model)
            print(
                f"  → 已加载标注样式：{LABEL_STYLE_JSON} "
                f"（{len(model['poly_offsets'])} 点，尖端=第 "
                f"{model.get('tip_index', ARROW_TIP_VERTEX_INDEX) + 1} 点）"
            )
            return model
        except Exception as exc:
            print(f"  → 读取 {LABEL_STYLE_JSON} 失败，使用内置默认：{exc}")

    set_rigid_label_model(build_canonical_rigid_model())
    print("  → 使用内置默认箭头+文字几何")
    return get_canonical_rigid()


def message_box(text: str, title: str = "引线标注") -> None:
    ctypes.windll.user32.MessageBoxW(0, text, title, 0)


# ---------- 与 xxq2022v01.py 相同的数据类型转换 ----------
def vtpnt(x, y, z=0.0):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(x), float(y), float(z)))


def vtfloat(lst):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, tuple(float(v) for v in lst))


def vtint(lst):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, tuple(int(v) for v in lst))


def vtvariant(lst):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, tuple(lst))


def list_open_documents(acad) -> list:
    """返回已打开图纸列表 [(Name, FullName, doc)]。"""
    docs = []
    try:
        count = int(acad.Documents.Count)
    except Exception:
        return docs
    for i in range(count):
        d = acad.Documents.Item(i)
        docs.append((str(d.Name), str(d.FullName), d))
    return docs


def activate_document_by_hint(acad, hint: str):
    """按文件名或路径关键字激活已打开的图纸（不新建、不 Open）。"""
    hint_norm = hint.replace("\\", "/").lower().strip()
    if not hint_norm:
        return acad.ActiveDocument
    for name, full, doc in list_open_documents(acad):
        name_l = name.lower()
        full_l = full.replace("\\", "/").lower()
        if (
            hint_norm == name_l
            or hint_norm == full_l
            or hint_norm in full_l
            or name_l == hint_norm
        ):
            doc.Activate()
            time.sleep(0.3)
            return doc
    raise RuntimeError(
        f"未在已打开图纸中找到「{hint}」。"
        f"请先在 CAD 中打开目标 DWG，或检查 --drawing 参数。"
    )


# 与 xxq2022v01.py 一致：版本后缀 → ProgID = AutoCAD.Application.{后缀}
CAD_VERSION_SUFFIX = {
    "2014": "19.1",
    "2016": "20",
    "2018": "22",
    "2020": "23.1",
    "2021": "24",
}
DEFAULT_CAD_VERSION = "22"  # AutoCAD 2018
AUTOCAD_PROGIDS = [
    "AutoCAD.Application.23",  # 2019
    "AutoCAD.Application.24",  # 2020
    "AutoCAD.Application.21",  # 2017
    "AutoCAD.Application.20",  # 2016
    "AutoCAD.Application.19",  # 2015
    "AutoCAD.Application",     # 通用（部分环境可用）
]


def _progid_sort_key(progid: str) -> int:
    """排序：优先 2018(.22)，其次其他带版本号 ProgID，最后无版本号。"""
    order = {p: i for i, p in enumerate(AUTOCAD_PROGIDS)}
    if progid in order:
        return order[progid]
    parts = progid.split(".")
    if len(parts) >= 3 and parts[-1].isdigit():
        return 50 - int(parts[-1])
    return 99


def _discover_autocad_progids() -> list:
    """从注册表扫描本机已注册的 AutoCAD.Application.* ProgID。"""
    found = []
    for i in range(16, 30):
        progid = f"AutoCAD.Application.{i}"
        try:
            winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid)
            found.append(progid)
        except OSError:
            pass
    try:
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "AutoCAD.Application")
        if "AutoCAD.Application" not in found:
            found.append("AutoCAD.Application")
    except OSError:
        pass
    found.sort(key=_progid_sort_key)
    return found


def _enumerate_autocad_from_rot() -> list:
    """从 ROT 枚举已运行的 AutoCAD（不启动新进程）。"""
    results = []
    try:
        ctx = pythoncom.CreateBindCtx(0)
        rot = pythoncom.GetRunningObjectTable()
        enum = rot.EnumRunning()
        while True:
            monikers = enum.Next(1)
            if not monikers:
                break
            moniker = monikers[0]
            name = moniker.GetDisplayName(ctx, None)
            if "AUTOCAD.APPLICATION" not in name.upper():
                continue
            obj = rot.GetObject(moniker)
            acad = Dispatch(obj)
            results.append((name, acad))
    except Exception:
        pass
    return results


def cad_connect(prog_id: str):
    """
    与 xxq2022v01.cad_connect 相同：Dispatch(prog_id)。
    AutoCAD 单实例模式下会绑到已打开的主程序，不会另开空白图。
    """
    wincad = Dispatch(prog_id)
    wincad.Visible = 1
    return wincad


def find_autocad_windows(keyword: str = None) -> list:
    """按窗口标题查找屏幕上的 AutoCAD 主窗口（含图纸名，如 测试1）。"""
    rows = []

    def cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        title = win32gui.GetWindowText(hwnd)
        if not title:
            return True
        tl = title.lower()
        if "autodesk" not in tl and "autocad" not in tl:
            return True
        if any(x in tl for x in ("mleader", "powershell", "cursor", "pycharm", "cmd.exe")):
            return True
        score = 20
        if "drawing1" in tl:
            score -= 80
        if keyword and keyword.lower() in title.lower():
            score += 300
        if "测试" in title or ".dwg" in title.lower():
            score += 40
        rows.append((score, hwnd, title))
        return True

    win32gui.EnumWindows(cb, None)
    rows.sort(key=lambda x: -x[0])
    return rows


def activate_best_document(acad, drawing_hint: str = None):
    """在同一 CAD 进程内激活得分最高的图纸（如 测试1.dwg）。"""
    candidates = []
    try:
        count = int(acad.Documents.Count)
    except Exception:
        count = 0
    for i in range(count):
        d = acad.Documents.Item(i)
        name = str(d.Name)
        full = str(d.FullName).strip()
        score = 0
        if full:
            score += 300
        if name.lower() != "drawing1.dwg":
            score += 100
        if drawing_hint:
            h = drawing_hint.replace("\\", "/").lower()
            nl = name.lower()
            fl = full.replace("\\", "/").lower()
            if h == nl or h in fl or nl == h or h in nl:
                score += 500
        candidates.append((score, d, name, full))
    if not candidates:
        raise RuntimeError("CAD 中没有打开的图纸")
    candidates.sort(key=lambda x: -x[0])
    print("  → 该 CAD 内图纸：")
    for score, _, name, full in candidates:
        print(f"      {name}  {full or '(未保存)'}")
    best = candidates[0][1]
    best.Activate()
    time.sleep(0.35)
    return acad.ActiveDocument


def is_blank_document(doc) -> bool:
    try:
        return (
            str(doc.Name).lower() == "drawing1.dwg"
            and not str(doc.FullName).strip()
        )
    except Exception:
        return True


def wait_user_open_drawing(acad, drawing_hint=None, open_path=None):
    """
    空白 Drawing1 或 HWND 未绑定时：弹窗让用户 文件→打开，或 --open 指定路径。
    """
    focus_autocad(acad)

    if open_path:
        if not os.path.isfile(open_path):
            raise FileNotFoundError(f"找不到 DWG：{open_path}")
        print(f"  → 正在打开：{open_path}")
        com_retry(lambda: acad.Documents.Open(open_path))
        time.sleep(0.8)
        doc = acad.ActiveDocument
        print(f"  → 已打开：{doc.Name}  ({doc.FullName})")
        return doc

    doc = acad.ActiveDocument
    if not is_blank_document(doc):
        if drawing_hint:
            doc = activate_best_document(acad, drawing_hint)
        return acad.ActiveDocument

    message_box(
        "脚本已连接到 AutoCAD。\n\n"
        "当前是空白图纸，请在该 AutoCAD 窗口中：\n"
        "  文件 → 打开 → 选择您的 DWG（如 测试1.dwg）\n\n"
        "打开完成后点击「确定」继续标注。\n\n"
        "也可下次运行加参数：\n"
        "  --open D:\\完整路径\\测试1.dwg",
        "请打开目标图纸",
    )
    focus_autocad(acad)
    doc = activate_best_document(acad, drawing_hint)
    if is_blank_document(doc):
        raise RuntimeError(
            "仍未打开目标图纸。请用 文件→打开 选择 DWG，"
            "或运行：python annotate_polylines.py --open 你的文件.dwg"
        )
    return acad.ActiveDocument


def _score_acad_instance(acad) -> tuple:
    """评分：有保存路径的图纸优先，避免连到脚本留下的空白 Drawing1。"""
    docs_info = []
    score = 0
    try:
        hwnd = int(acad.HWND)
    except Exception:
        hwnd = 0
    try:
        for name, full, _ in list_open_documents(acad):
            full_s = str(full).strip()
            docs_info.append((name, full_s))
            if full_s:
                score += 300
            if str(name).lower() != "drawing1.dwg":
                score += 100
            score += 10
    except Exception:
        pass
    return score, docs_info, hwnd


def _gather_running_autocad(progids: list) -> list:
    """仅 ROT + GetActiveObject，不 Dispatch（避免先开出空白 CAD）。"""
    candidates = []
    seen_hwnd = set()

    def add(acad, method, label):
        try:
            hwnd = int(acad.HWND)
        except Exception:
            return
        if hwnd in seen_hwnd:
            return
        seen_hwnd.add(hwnd)
        score, docs_info, hwnd = _score_acad_instance(acad)
        candidates.append({
            "label": label,
            "acad": acad,
            "method": method,
            "score": score,
            "docs": docs_info,
            "hwnd": hwnd,
        })

    for name, acad in _enumerate_autocad_from_rot():
        add(acad, "ROT", name)

    for progid in progids:
        try:
            add(win32com.client.GetActiveObject(progid), "GetActiveObject", progid)
        except Exception:
            pass

    candidates.sort(key=lambda x: -x["score"])
    return candidates


def _print_autocad_instances(instances: list) -> None:
    print(f"  → 检测到 {len(instances)} 个 AutoCAD COM 实例：")
    for i, inst in enumerate(instances):
        print(
            f"      [{i}] HWND={inst['hwnd']} score={inst['score']} "
            f"({inst['method']})"
        )
        for dn, df in inst["docs"]:
            path = df if df else "(未保存)"
            print(f"           {dn}  {path}")


def _pick_autocad_candidate(
    candidates: list,
    instance_index: int = None,
    drawing_hint: str = None,
    windows: list = None,
) -> dict:
    if not candidates:
        return None

    if instance_index is not None:
        if instance_index < 0 or instance_index >= len(candidates):
            raise ValueError(
                f"--cad-instance {instance_index} 无效，"
                f"共 {len(candidates)} 个（0～{len(candidates)-1}）"
            )
        return candidates[instance_index]

    if windows:
        target_hwnd = windows[0][1]
        for c in candidates:
            if c["hwnd"] == target_hwnd:
                print(f"  → 已按屏幕窗口 HWND={target_hwnd} 匹配 COM")
                return c

    if drawing_hint:
        hint = drawing_hint.replace("\\", "/").lower()
        for c in candidates:
            for dn, df in c["docs"]:
                dn_l = dn.lower()
                df_l = df.replace("\\", "/").lower()
                if hint == dn_l or hint in df_l or hint in dn_l:
                    return c

    candidates.sort(key=lambda x: -x["score"])
    pick = candidates[0]
    if windows and pick["hwnd"] != windows[0][1]:
        print(
            f"  → 提示：COM HWND={pick['hwnd']} ≠ 屏幕窗口 HWND={windows[0][1]}\n"
            f"      若图纸不对：python annotate_polylines.py --cad-instance 1"
        )
    return pick


def _build_progids(cad_version: str) -> list:
    """生成待尝试的 ProgID 列表（优先用户指定版本，与参考脚本 Dispatch 方式一致）。"""
    ver = str(cad_version).strip()
    if ver in CAD_VERSION_SUFFIX:
        ver = CAD_VERSION_SUFFIX[ver]
    primary = f"AutoCAD.Application.{ver}"
    progids = [primary]
    for p in _discover_autocad_progids():
        if p not in progids:
            progids.append(p)
    for p in AUTOCAD_PROGIDS:
        if p not in progids:
            progids.append(p)
    return progids


def connect_autocad(
    drawing_hint: str = None,
    cad_version: str = DEFAULT_CAD_VERSION,
    allow_new_cad: bool = False,
    instance_index: int = None,
    open_path: str = None,
):
    """
    连接 AutoCAD：先尝试绑定屏幕窗口；否则 Dispatch(xxq) 并等待用户打开 DWG。
    """
    progids = _build_progids(cad_version)
    prog_id = progids[0]

    windows = find_autocad_windows(drawing_hint)
    if windows:
        print("  → 屏幕 AutoCAD 窗口：")
        for score, hwnd, title in windows[:6]:
            print(f"      HWND={hwnd} score={score}  {title}")

    candidates = _gather_running_autocad(progids)
    if candidates:
        _print_autocad_instances(candidates)

    target_hwnd = windows[0][1] if windows else None
    pick = _pick_autocad_candidate(
        candidates, instance_index, drawing_hint, windows,
    )

    acad = None
    need_dispatch = False

    if pick and instance_index is not None:
        acad = pick["acad"]
        print(f"  → 使用 --cad-instance {instance_index}")
    elif pick and target_hwnd and pick["hwnd"] == target_hwnd:
        acad = pick["acad"]
        print(f"  → 已绑定屏幕窗口 HWND={target_hwnd}")
    elif pick and not target_hwnd:
        acad = pick["acad"]
    elif pick and target_hwnd and pick["hwnd"] != target_hwnd:
        print(
            f"  → 屏幕窗口 HWND={target_hwnd}，"
            f"COM HWND={pick['hwnd']} 不一致，将 Dispatch 连接"
        )
        need_dispatch = True
    elif not candidates:
        need_dispatch = True

    if acad is None or need_dispatch or allow_new_cad:
        if acad is None or need_dispatch:
            message_box(
                "将通过 Dispatch 连接 AutoCAD（与 xxq 脚本相同）。\n\n"
                "若出现新的空白窗口，请使用 文件→打开 选择目标 DWG，\n"
                "或命令行加：--open D:\\路径\\测试1.dwg\n\n"
                "点击确定继续。",
                "连接 AutoCAD",
            )
        acad = cad_connect(prog_id)
        print(f"  → Dispatch 已连接（HWND={int(acad.HWND)}）")

    focus_autocad(acad)
    try:
        ver = acad.Version
    except Exception:
        ver = "?"
    print(f"  → AutoCAD Version={ver}")

    doc = wait_user_open_drawing(acad, drawing_hint, open_path)
    print(f"  → 当前图纸：{doc.Name}  ({doc.FullName or '未保存'})")
    return acad, doc


def _delete_selection_set(acad, name: str) -> None:
    try:
        acad.ActiveDocument.SelectionSets.Item(name).Delete()
    except Exception:
        pass


def focus_autocad(acad) -> None:
    """把键盘焦点切到 AutoCAD，减少 COM「被呼叫方拒绝接收呼叫」。"""
    try:
        hwnd = int(acad.HWND)
        ctypes.windll.user32.SetForegroundWindow(hwnd)
    except Exception:
        pass
    try:
        acad.WindowState = 1
    except Exception:
        pass
    time.sleep(0.35)


def com_retry(fn, retries: int = 8, delay: float = 0.4):
    """COM 忙时重试（RPC_E_CALL_REJECTED / 被呼叫方拒绝接收呼叫）。"""
    last_err = None
    for _ in range(retries):
        try:
            return fn()
        except Exception as exc:
            last_err = exc
            err = str(exc)
            code = getattr(exc, "args", [None])[0]
            if (
                code == -2147418111
                or "-2147418111" in err
                or "拒绝" in err
                or "rejected" in err.lower()
                or "call rejected" in err.lower()
            ):
                time.sleep(delay)
                continue
            raise
    raise last_err


def merge_style_with_defaults(info: dict) -> dict:
    """图中读到的样式可能与默认合并（空 style 名等）。"""
    merged = dict(DEFAULT_STYLE_INFO)
    for key, val in info.items():
        if val is None:
            continue
        if isinstance(val, str) and not val.strip():
            continue
        merged[key] = val
    return merged


def _make_entity_filter(object_names: frozenset):
    """SelectOnScreen 过滤器（VARIANT 格式，与 xxq 脚本一致）。"""
    names = sorted(object_names)
    if len(names) == 1:
        return vtint([0]), vtvariant([names[0]])
    types = [-4] + [0] * len(names) + [-4]
    data = ["<OR"] + list(names) + ["OR>"]
    return vtint(types), vtvariant(data)


def read_id_list(excel_path: str) -> list:
    """读取 Excel 第一列 ID，跳过表头行（若首格为 ID/编号 等）。"""
    if not os.path.isfile(excel_path):
        raise FileNotFoundError(f"找不到 Excel：{excel_path}")
    df = pd.read_excel(excel_path, header=None, usecols=[0])
    ids = []
    for val in df.iloc[:, 0]:
        if pd.isna(val):
            continue
        s = str(val).strip()
        if not s:
            continue
        if not ids and s.lower() in ("id", "编号", "name", "名称", "序号"):
            continue
        ids.append(s)
    if not ids:
        raise ValueError(f"Excel 第一列未读到有效 ID：{excel_path}")
    return ids


def pick_entity(
    acad, prompt: str, allowed_names: frozenset, use_filter: bool = False,
):
    """
    屏幕点选。默认不用过滤器（AcDbMLeader 等过滤器常导致选不上）。
    """
    allowed_str = "、".join(sorted(allowed_names))
    focus_autocad(acad)
    print(f"  → {prompt}（请在 CAD 图中点选，Esc 取消）")

    doc = acad.ActiveDocument
    sset_name = f"PyPick_{int(time.time() * 1000)}"
    _delete_selection_set(acad, sset_name)
    sset = doc.SelectionSets.Add(sset_name)
    ft, fd = _make_entity_filter(allowed_names)

    while True:
        try:
            if use_filter:
                com_retry(
                    lambda: sset.SelectOnScreen(ft, fd),
                    retries=15, delay=0.5,
                )
            else:
                com_retry(
                    lambda: sset.SelectOnScreen(),
                    retries=15, delay=0.5,
                )
        except Exception as exc:
            err = str(exc).lower()
            if any(k in err for k in ("cancel", "keyword", "用户", "拒绝", "reject")):
                try:
                    sset.Delete()
                except Exception:
                    pass
                raise KeyboardInterrupt("用户取消选择")
            try:
                com_retry(lambda: sset.SelectOnScreen(), retries=15, delay=0.5)
            except Exception as exc2:
                try:
                    sset.Delete()
                except Exception:
                    pass
                err2 = str(exc2).lower()
                if any(k in err2 for k in ("cancel", "keyword", "用户", "拒绝")):
                    raise KeyboardInterrupt("用户取消选择")
                raise exc2

        # SelectOnScreen 返回后 CAD 可能仍忙，需等待并重试读 Count
        time.sleep(0.25)
        try:
            count = com_retry(lambda: int(sset.Count), retries=20, delay=0.5)
        except Exception as exc:
            err = str(exc).lower()
            if any(k in err for k in ("cancel", "keyword", "用户")):
                raise KeyboardInterrupt("用户取消选择")
            print(f"  → 读取选择集失败，请重新点选… ({exc})")
            focus_autocad(acad)
            continue

        if count == 0:
            print(f"  → 未选中对象，需要：{allowed_str}，请重试…")
            focus_autocad(acad)
            continue

        entity = com_retry(lambda: sset.Item(0), retries=12, delay=0.4)
        name = com_retry(lambda: entity.ObjectName, retries=12, delay=0.4)
        if name not in allowed_names:
            print(f"  → 选中 [{name}]，需要 {allowed_str}，请重新选择…")
            try:
                sset.Delete()
            except Exception:
                pass
            sset_name = f"PyPick_{int(time.time() * 1000)}"
            doc = acad.ActiveDocument
            sset = doc.SelectionSets.Add(sset_name)
            focus_autocad(acad)
            continue

        try:
            sset.Delete()
        except Exception:
            pass
        return entity, None


def get_polyline_vertices(entity) -> list:
    """提取二维多段线 / 轻量多段线各顶点 (x, y)。"""
    name = entity.ObjectName
    verts = []

    if name == "AcDbPolyline":
        coords = list(entity.Coordinates)
        for i in range(0, len(coords), 2):
            verts.append((float(coords[i]), float(coords[i + 1])))
    elif name == "AcDb2dPolyline":
        n = int(entity.NumberOfVertices)
        for i in range(n):
            c = entity.Coordinate(i)
            verts.append((float(c[0]), float(c[1])))
    else:
        raise ValueError(f"不支持的多段线类型：{name}")

    if len(verts) < 1:
        raise ValueError("多段线没有顶点")
    return verts


def read_mleader_style(ref_mleader) -> dict:
    """从参考多重引线读取样式名、比例及常用显示参数。"""
    info = {}
    try:
        info["style"] = str(ref_mleader.Style)
    except Exception:
        info["style"] = ""
    try:
        info["scale"] = float(ref_mleader.ScaleFactor)
    except Exception:
        info["scale"] = 1.0
    try:
        info["text_height"] = float(ref_mleader.TextHeight)
    except Exception:
        info["text_height"] = None
    try:
        info["arrowhead_size"] = float(ref_mleader.ArrowheadSize)
    except Exception:
        info["arrowhead_size"] = None
    try:
        info["landing_distance"] = float(ref_mleader.DoglegLength)
    except Exception:
        try:
            info["landing_distance"] = float(ref_mleader.LandingDistance)
        except Exception:
            info["landing_distance"] = None
    try:
        info["text_landing_gap"] = float(ref_mleader.LandingGap)
    except Exception:
        info["text_landing_gap"] = None
    try:
        info["text_style"] = str(ref_mleader.TextStyleName)
    except Exception:
        try:
            info["text_style"] = str(ref_mleader.MText.StyleName)
        except Exception:
            info["text_style"] = None
    try:
        info["content_type"] = int(ref_mleader.ContentType)
    except Exception:
        info["content_type"] = None
    # 兼容旧字段名
    if info.get("landing_distance") is not None:
        info["landing_gap"] = info["landing_distance"]
    return info


def effective_style_dims(style_info: dict, label: str = "S3-03") -> dict:
    """名义尺寸 × 全局比例 → 图面实际尺寸（例：字高 10×0.015=0.15）。"""
    scale = float(style_info.get("scale", 0.015) or 0.015)
    text_nom = float(style_info.get("text_height", 10) or 10)
    arrow_nom = float(style_info.get("arrowhead_size", 10) or 10)
    land_nom = float(
        style_info.get("landing_distance")
        or style_info.get("landing_gap")
        or 1.44
    )
    gap_nom = float(style_info.get("text_landing_gap", 8) or 8)
    if style_info.get("drawing_text_height") is not None:
        eff_h = float(style_info["drawing_text_height"])
    else:
        eff_h = text_nom * scale
    eff_w = max(len(label) * eff_h * 0.65, eff_h * 2.0)
    return {
        "scale": scale,
        "text_height": eff_h,
        "arrow": arrow_nom * scale,
        "landing_dist": land_nom * scale,
        "landing_gap": gap_nom * scale,
        "text_width": eff_w,
    }


def find_mleader_style_in_model(acad) -> Optional[dict]:
    """若图里已有多重引线，自动读取其样式（无需手选）。"""
    try:
        ms = acad.ActiveDocument.ModelSpace
        count = int(ms.Count)
        for i in range(count):
            ent = ms.Item(i)
            if ent.ObjectName == MLEADER_OBJECT_NAME:
                info = merge_style_with_defaults(read_mleader_style(ent))
                print(
                    f"  → 从图中已有引线读取样式: {info.get('style')!r}, "
                    f"比例={info.get('scale')}"
                )
                return info
    except Exception:
        pass
    return None


def resolve_style_info(acad, pick_reference: bool = False) -> dict:
    """样式：图中已有引线 → 默认 CBDI → 可选手选参考。"""
    info = find_mleader_style_in_model(acad)
    if info:
        return info

    info = dict(DEFAULT_STYLE_INFO)
    print(
        f"  → 使用默认引线样式: {info['style']!r}, "
        f"比例={info['scale']}, 字高={info['text_height']}"
    )

    if pick_reference:
        try:
            print("  → 可选：点选一条参考多重引线复制样式（Esc 跳过）")
            ref, _ = pick_entity(
                acad, "点选参考多重引线", frozenset({MLEADER_OBJECT_NAME}),
                use_filter=False,
            )
            info = read_mleader_style(ref)
            info = merge_style_with_defaults(info)
            print(f"  → 已用手选参考样式: {info.get('style')!r}")
        except KeyboardInterrupt:
            print("  → 未选手选参考，继续用默认样式")

    return info


class LabelPlacementContext:
    """记录已放置标注的包围盒，用于避让已有标注。"""

    def __init__(self, style_info: dict):
        self.style_info = style_info
        self.obstacles: list = []

    def add_box(self, cx: float, cy: float, hw: float, hh: float) -> None:
        self.obstacles.append((cx, cy, hw, hh))

    def overlap_score(
        self, cx: float, cy: float, hw: float, hh: float, margin: float = 0.0,
    ) -> float:
        score = 0.0
        for ox, oy, ohw, ohh in self.obstacles:
            dx = abs(cx - ox)
            dy = abs(cy - oy)
            ox_ = max(0.0, (hw + ohw + margin) - dx)
            oy_ = max(0.0, (hh + ohh + margin) - dy)
            score += ox_ * oy_
        return score


def collect_existing_obstacles(acad, ctx: LabelPlacementContext) -> None:
    """扫描图中已有多重引线/文字，作为避让障碍。"""
    dims = effective_style_dims(ctx.style_info)
    default_h = dims["text_height"]
    try:
        ms = acad.ActiveDocument.ModelSpace
        count = int(ms.Count)
    except Exception:
        return

    for i in range(count):
        try:
            ent = ms.Item(i)
            name = ent.ObjectName
            if name == MLEADER_OBJECT_NAME:
                try:
                    p = ent.TextLocation
                    cx, cy = float(p[0]), float(p[1])
                    try:
                        txt = str(ent.TextString or "")
                    except Exception:
                        txt = ""
                    w = max(len(txt) * default_h * 0.65, default_h * 2.5)
                    ctx.add_box(cx, cy, w * 0.5, default_h * 0.85)
                except Exception:
                    pass
            elif name == "AcDbMText":
                try:
                    p = ent.InsertionPoint
                    cx, cy = float(p[0]), float(p[1])
                    h = float(ent.Height)
                    try:
                        w = float(ent.Width)
                    except Exception:
                        w = len(str(ent.TextString or "")) * h * 0.65
                    ctx.add_box(cx, cy, max(w, h) * 0.5, h * 0.85)
                except Exception:
                    pass
            elif name == "AcDbText":
                try:
                    p = ent.InsertionPoint
                    cx, cy = float(p[0]), float(p[1])
                    h = float(ent.Height)
                    w = len(str(ent.TextString or "")) * h * 0.65
                    ctx.add_box(cx, cy, max(w, h), h * 0.85)
                except Exception:
                    pass
        except Exception:
            pass


def vertex_geometry(vx, vy, prev_pt, next_pt):
    """顶点处切向与法向（法向用于把文字推向多段线外侧）。"""
    if prev_pt is not None and next_pt is not None:
        v1x, v1y = vx - prev_pt[0], vy - prev_pt[1]
        v2x, v2y = next_pt[0] - vx, next_pt[1] - vy
        l1 = math.hypot(v1x, v1y) or 1.0
        l2 = math.hypot(v2x, v2y) or 1.0
        tx = v1x / l1 + v2x / l2
        ty = v1y / l1 + v2y / l2
        tl = math.hypot(tx, ty) or 1.0
        tx, ty = tx / tl, ty / tl
    elif next_pt is not None:
        dx, dy = next_pt[0] - vx, next_pt[1] - vy
        l = math.hypot(dx, dy) or 1.0
        tx, ty = dx / l, dy / l
    elif prev_pt is not None:
        dx, dy = vx - prev_pt[0], vy - prev_pt[1]
        l = math.hypot(dx, dy) or 1.0
        tx, ty = dx / l, dy / l
    else:
        tx, ty = 1.0, 0.0
    nx, ny = -ty, tx
    return tx, ty, nx, ny


def annotation_segment_direction(vx, vy, prev_pt, next_pt) -> tuple:
    """
    标注引线末段应与之平行的多段线段方向。
    拐弯点：仅用前一段（prev→当前顶点）；无端点则用下一段。
    """
    if prev_pt is not None:
        dx = vx - prev_pt[0]
        dy = vy - prev_pt[1]
    elif next_pt is not None:
        dx = next_pt[0] - vx
        dy = next_pt[1] - vy
    else:
        return 1.0, 0.0
    length = math.hypot(dx, dy) or 1.0
    return dx / length, dy / length


def build_label_at_tip(
    vx: float, vy: float,
    u_x: float, u_y: float,
) -> tuple:
    """
    刚性整体绕尖端旋转：使 u（P6→P7 / 文字方向）对齐目标方向。
    返回 (flat, text_x, text_y, text_rot_rad, u_x, u_y)。
    """
    model = get_canonical_rigid()
    ux, uy = _unit(u_x, u_y)
    delta = math.atan2(uy, ux) - model["ang_can"]
    cos_r = math.cos(delta)
    sin_r = math.sin(delta)

    flat = []
    for ox, oy in model["poly_offsets"]:
        rx, ry = _rot2d(ox, oy, cos_r, sin_r)
        flat.extend([vx + rx, vy + ry])

    tox, toy = model["text_offset"]
    text_x, text_y = _rot2d(tox, toy, cos_r, sin_r)
    text_x += vx
    text_y += vy
    text_rot = model["text_rot"] + delta

    return flat, text_x, text_y, text_rot, ux, uy


def _parallel_angle_distance(ang: float, base_ang: float) -> float:
    """与 base 或 base+π 的最小角度差（弧度）。"""
    d = abs((ang - base_ang + math.pi) % (2 * math.pi) - math.pi)
    d_rev = abs((ang - base_ang - math.pi + math.pi) % (2 * math.pi) - math.pi)
    return min(d, d_rev)


def choose_leader_geometry(
    vx: float, vy: float,
    prev_pt, next_pt,
    style_info: dict,
    label: str,
    ctx: LabelPlacementContext,
) -> tuple:
    """
    默认：整体 u 方向与多段线段平行；避障时绕尖端旋转整体（可不平行）。
    返回 (flat, text_x, text_y, text_rot_rad, u_x, u_y)。
    """
    base_seg_x, base_seg_y = annotation_segment_direction(vx, vy, prev_pt, next_pt)
    base_ang = math.atan2(base_seg_y, base_seg_x)
    dims = effective_style_dims(style_info, label)
    eff_h = dims["text_height"]
    eff_w = dims["text_width"]
    margin = eff_h * 0.4
    hw = eff_w * 0.52
    hh = eff_h * 0.85

    candidates = []

    # 优先：与多段线段平行（±方向）
    for flip in (0, 1):
        ang = base_ang + (math.pi if flip else 0.0)
        sx, sy = math.cos(ang), math.sin(ang)
        geom = build_label_at_tip(vx, vy, sx, sy)
        tx, ty = geom[1], geom[2]
        overlap = ctx.overlap_score(tx, ty, hw, hh, margin)
        candidates.append((overlap, 0.0, geom))

    min_parallel_overlap = min(c[0] for c in candidates)

    # 有遮挡时：绕尖端整体旋转（步长 30°），允许不平行
    if min_parallel_overlap > 0:
        for deg in range(0, 360, AVOIDANCE_ROT_STEP_DEG):
            ang = base_ang + math.radians(deg)
            if _parallel_angle_distance(ang, base_ang) < 0.05:
                continue
            sx, sy = math.cos(ang), math.sin(ang)
            geom = build_label_at_tip(vx, vy, sx, sy)
            tx, ty = geom[1], geom[2]
            overlap = ctx.overlap_score(tx, ty, hw, hh, margin)
            parallel_dist = _parallel_angle_distance(ang, base_ang)
            candidates.append((overlap, parallel_dist * 0.4, geom))

    candidates.sort(key=lambda c: c[0] + c[1])
    flat, text_x, text_y, text_rot, ux, uy = candidates[0][2]
    ctx.add_box(text_x, text_y, hw, hh)
    return flat, text_x, text_y, text_rot, ux, uy


def set_entity_layer(entity, layer: Optional[str]) -> None:
    if not layer:
        return
    try:
        entity.Layer = layer
    except Exception:
        pass


def create_manual_label(
    ms,
    vx: float, vy: float,
    label: str,
    style_info: dict,
    geometry: tuple,
    layer: Optional[str] = None,
):
    """
    刚性整体：等边箭头多段线 + 文字（相对位置固定），绕尖端对齐待标注多段线。
    """
    flat, text_x, text_y, text_rot, _ux, _uy = geometry
    dims = effective_style_dims(style_info, label)
    rigid = get_canonical_rigid()
    eff_h = dims["text_height"]
    if rigid.get("text_height") is not None:
        eff_h = float(rigid["text_height"])
    eff_w = dims["text_width"]

    leader_pl = ms.AddLightWeightPolyline(vtfloat(flat))

    mt = ms.AddMText(vtpnt(text_x, text_y), eff_w, label)
    mt.Height = eff_h
    attach = rigid.get("text_attachment")
    try:
        mt.AttachmentPoint = attach if attach is not None else 5
    except Exception:
        pass
    try:
        mt.Rotation = text_rot
    except Exception:
        pass
    try:
        ts = style_info.get("text_style")
        if ts:
            mt.StyleName = ts
    except Exception:
        pass

    set_entity_layer(leader_pl, layer)
    set_entity_layer(mt, layer)

    return mt


def create_point_label(
    ms, vx, vy, label, style_info, geometry, layer=None,
):
    return create_manual_label(ms, vx, vy, label, style_info, geometry, layer)


def annotate_polyline(
    acad, entity, section_id: str, style_info: dict, ctx: LabelPlacementContext,
) -> list:
    """对一条多段线所有顶点创建引线标注，返回坐标记录行。"""
    doc = acad.ActiveDocument
    verts = get_polyline_vertices(entity)
    ms = doc.ModelSpace
    try:
        poly_layer = str(entity.Layer)
    except Exception:
        poly_layer = None
    records = []
    total = len(verts)

    for i, (vx, vy) in enumerate(verts):
        prev_pt = verts[i - 1] if i > 0 else None
        next_pt = verts[i + 1] if i < total - 1 else None
        label = f"{section_id}-{i + 1:02d}"
        geometry = choose_leader_geometry(
            vx, vy, prev_pt, next_pt, style_info, label, ctx,
        )
        create_point_label(
            ms, vx, vy, label, style_info, geometry, poly_layer,
        )
        records.append({
            "ID": section_id,
            "标注": label,
            "点序号": i + 1,
            "X": round(vx, 4),
            "Y": round(vy, 4),
        })

    try:
        doc.Regen(1)  # acActiveViewport
    except Exception:
        pass
    return records


def save_output(records: list, path: str) -> None:
    df = pd.DataFrame(records)
    if df.empty:
        df = pd.DataFrame(columns=["ID", "标注", "点序号", "X", "Y"])

    base, ext = os.path.splitext(path)
    fallbacks = [path]
    if ext.lower() == ".xlsx":
        fallbacks.append(f"{base}_backup.xlsx")
    fallbacks.append(os.path.join(SCRIPT_DIR, f"output_{int(time.time())}.xlsx"))

    last_err = None
    for target in fallbacks:
        try:
            df.to_excel(target, index=False)
            print(f"  → 坐标已保存：{target}（{len(df)} 行）")
            if target != path:
                print(
                    f"  → 提示：无法写入 {path}（文件可能被 Excel 占用），"
                    f"已写入备用文件。请关闭 Excel 后重试。"
                )
            return
        except PermissionError as exc:
            last_err = exc
            continue
        except OSError as exc:
            if exc.errno == 13:
                last_err = exc
                continue
            raise

    raise PermissionError(
        f"无法写入 Excel：{path}。请关闭已打开的 output.xlsx 后重试。"
    ) from last_err


def run(
    excel_path: str,
    drawing_hint: str = None,
    cad_version: str = DEFAULT_CAD_VERSION,
    allow_new_cad: bool = False,
    instance_index: int = None,
    open_path: str = None,
    pick_style: bool = False,
    pick_label: bool = False,
) -> None:
    pythoncom.CoInitialize()
    try:
        ids = read_id_list(excel_path)
        print(f"  → 从 Excel 读取 {len(ids)} 个 ID：{', '.join(ids)}")

        acad, doc = connect_autocad(
            drawing_hint, cad_version, allow_new_cad, instance_index, open_path,
        )

        style_info = resolve_style_info(acad, pick_reference=pick_style)
        rigid_model = resolve_rigid_label_style(acad, pick_label=pick_label)
        if rigid_model:
            style_info = merge_style_from_rigid_model(style_info, rigid_model)

        dims = effective_style_dims(style_info)
        print(
            f"  → 标注参数: 全局比例={style_info.get('scale')}, "
            f"图面字高≈{dims['text_height']:.4f}, 箭头≈{dims['arrow']:.4f}"
        )

        placement_ctx = LabelPlacementContext(style_info)
        collect_existing_obstacles(acad, placement_ctx)
        print(f"  → 已扫描图中 {len(placement_ctx.obstacles)} 处已有标注用于避让")

        all_records = []
        if os.path.isfile(OUTPUT_EXCEL):
            try:
                old = pd.read_excel(OUTPUT_EXCEL)
                all_records = old.to_dict("records")
            except Exception:
                pass

        for idx, section_id in enumerate(ids):
            if idx > 0:
                next_hint = ids[idx] if idx < len(ids) else ""
                message_box(
                    f"已完成 ID「{ids[idx - 1]}」的标注。\n\n"
                    f"点击确定后，请在图中选择 ID「{section_id}」对应的多段线。",
                    "继续标注",
                )

            print(f"  → 请选择 ID「{section_id}」对应的二维多段线")
            poly, _ = pick_entity(
                acad,
                f"请选择 [{section_id}] 的多段线",
                POLYLINE_OBJECT_NAMES,
            )

            print(f"  → 正在标注 {section_id} …")
            rows = annotate_polyline(acad, poly, section_id, style_info, placement_ctx)
            all_records.extend(rows)
            save_output(all_records, OUTPUT_EXCEL)
            print(f"  → {section_id} 完成，共 {len(rows)} 个顶点")

        message_box(
            f"全部完成！共处理 {len(ids)} 个 ID。\n坐标文件：{OUTPUT_EXCEL}",
            "完成",
        )
    finally:
        pythoncom.CoUninitialize()


def main():
    parser = argparse.ArgumentParser(description="AutoCAD 多段线顶点引线标注（箭头+多段线+文字）")
    parser.add_argument(
        "--excel", "-e",
        default=DEFAULT_EXCEL,
        help=f"ID 列表 Excel（第一列），默认 {DEFAULT_EXCEL}",
    )
    parser.add_argument(
        "--drawing", "-d",
        default=None,
        help="指定已打开图纸的文件名或路径关键字",
    )
    parser.add_argument(
        "--cad-version", "-c",
        default=DEFAULT_CAD_VERSION,
        help="AutoCAD 版本后缀或年份，2018 对应 22，默认 22",
    )
    parser.add_argument(
        "--open", "-o",
        default=None,
        help="直接打开 DWG 完整路径（新窗口时推荐）",
    )
    parser.add_argument(
        "--allow-new-cad",
        action="store_true",
        help="找不到已运行 CAD 时允许 Dispatch 启动新实例（默认禁止）",
    )
    parser.add_argument(
        "--cad-instance", "-i",
        type=int,
        default=None,
        help="指定 CAD 进程序号（先运行一次看 [0][1] 列表）",
    )
    parser.add_argument(
        "--list-cad",
        action="store_true",
        help="仅列出已运行的 AutoCAD 进程后退出",
    )
    parser.add_argument(
        "--pick-style",
        action="store_true",
        help="可选手选参考多重引线，读取字高/比例等（与 --pick-label 可同时使用）",
    )
    parser.add_argument(
        "--pick-label",
        action="store_true",
        help="点选图中箭头多段线+文字，记录样式到 label_style.json（尖端=第3点）",
    )
    args = parser.parse_args()

    if args.list_cad:
        pythoncom.CoInitialize()
        try:
            progids = _build_progids(args.cad_version)
            instances = _gather_running_autocad(progids)
            windows = find_autocad_windows(args.drawing)
            if windows:
                print("  → 屏幕 AutoCAD 窗口：")
                for score, hwnd, title in windows[:6]:
                    print(f"      HWND={hwnd}  {title}")
            if not instances:
                print("  → 未检测到运行中的 AutoCAD")
            else:
                _print_autocad_instances(instances)
        finally:
            pythoncom.CoUninitialize()
        return

    try:
        run(
            args.excel, args.drawing, args.cad_version,
            args.allow_new_cad, args.cad_instance, args.open,
            args.pick_style, args.pick_label,
        )
    except KeyboardInterrupt:
        print("\n  → 用户取消")
        sys.exit(0)
    except Exception as exc:
        print(f"\n  ✗ 错误：{exc}")
        traceback.print_exc()
        message_box(f"运行出错：\n{exc}", "错误")
        sys.exit(1)


if __name__ == "__main__":
    main()
