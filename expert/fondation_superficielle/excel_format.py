# -*- coding: utf-8 -*-
"""输入/输出 Excel 分区配色与计算书格式输出"""

from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from midas_import import MIDAS_COLUMNS, MIDAS_SECTION_KEY

FILL_TITLE = PatternFill("solid", fgColor="2F5496")
FILL_SECTION_GEOM = PatternFill("solid", fgColor="C6E0B4")
FILL_SECTION_GEOTECH = PatternFill("solid", fgColor="F8CBAD")
FILL_SECTION_SLIDE = PatternFill("solid", fgColor="FFE699")
FILL_SECTION_SETTLE = PatternFill("solid", fgColor="BDD7EE")
FILL_SECTION_REINF = PatternFill("solid", fgColor="E2EFDA")
FILL_SECTION_LOAD = PatternFill("solid", fgColor="D9E1F2")
FILL_HEADER = PatternFill("solid", fgColor="D9E1F2")
FILL_DESC = PatternFill("solid", fgColor="F2F2F2")
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")
FILL_OUTPUT_TITLE = PatternFill("solid", fgColor="44546A")
FILL_OUTPUT_SECTION = PatternFill("solid", fgColor="ADB9CA")

FONT_TITLE = Font(bold=True, size=14, color="FFFFFF")
FONT_SECTION = Font(bold=True, size=11)
FONT_HEADER = Font(bold=True, size=10)
FONT_NORMAL = Font(size=10)
FONT_OUTPUT_TITLE = Font(bold=True, size=12, color="FFFFFF")
FONT_OUTPUT_SECTION = Font(bold=True, size=11, color="FFFFFF")

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
THIN = Side(style="thin", color="B4B4B4")
BORDER_ALL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

KNOWN_COL_KEYS = {
    "N_kN", "ex_m", "ey_m", "Hd_kN", "section",
    "Mmax_ELU_kNm_m", "Mmax_ELS_kNm_m", "Vd_ELU_kN",
    "no", "depth_m", "lithology", "pf_MPa", "plm_MPa", "EM_MPa", "E_over_plm",
}

LOAD_SECTION_KEYS = (
    "ELU_mobilisation", "ELS_mobilisation", "ELU_glissement", "Ferraillage",
    MIDAS_SECTION_KEY,
)
PRESSIOMETER_SECTION_KEY = "Pressiometre_Menard"
ALL_SECTION_KEYS = (PRESSIOMETER_SECTION_KEY,) + LOAD_SECTION_KEYS
PIER_PARAM_SECTIONS: List[Tuple[str, PatternFill, List[Tuple[str, Any, str, str]]]] = [
    (
        "【一、本墩基础几何尺寸】",
        FILL_SECTION_GEOM,
        [
            ("H_cm", 130, "承台厚度 h", "cm"),
            ("B_cm", 900, "横桥向宽度（÷ 此值得单位宽度）", "cm"),
            ("L_cm", 920, "纵桥向长度（q 分布方向；B₁/B₂ 沿此向）", "cm"),
            ("De_cm", 165, "基础埋深 De（地面→基底）", "cm"),
            ("Zw_cm", 0, "地下水位埋深（自地面起算）", "cm"),
            ("midas_element_ids", "", "Midas 承台单元（总 N、My；逗号分隔）", "—"),
            ("pier_a_cm", "", "墩纵桥向宽度 a（居中→B₁=B₂=(L−a)/2）", "cm"),
            ("B_1_cm", "", "纵桥向 A-A 至 q_max 外缘 B₁", "cm"),
            ("B_2_cm", "", "纵桥向 B-B 至 q_min 外缘 B₂", "cm"),
        ],
    ),
    (
        "【二、本墩土工与承载力】",
        FILL_SECTION_GEOTECH,
        [
            ("gamma_kN_m3", 20, "天然重度 γ（水位以上）", "kN/m³"),
            ("gamma_buoy_kN_m3", "", "浮重度 γ′（水位以下；留空=γ−10）", "kN/m³"),
            ("q0_kPa", "", "q′₀ (kPa)；留空则 γ×De 自动算", "kPa"),
            ("kp", 1.02, "承载力系数 kp（EC7/Fasc.62 特征值）", "—"),
            ("ple_star_kPa", "", "ple* (kPa)；留空则由旁压试验计算", "kPa"),
            ("q_u_kPa", "", "可选：直接填 q′u，留空则 q₀+kp×ple*", "kPa"),
            ("phi_deg", 25, "内摩擦角 φ", "°"),
            ("C_kPa", 75, "有效粘聚力 C'", "kPa"),
            ("soil_category", "", "土质（argile/sable/mixte；留空自动判定）", "—"),
            ("replacement_thickness_m", "", "换填厚度（自基底向下，m）", "m"),
            ("replacement_plm_MPa", "", "换填材料 D2/D3 旁压 plm", "MPa"),
            ("ple_star_zone", "replacement", "ple*：mixed=1.5B混合；replacement=仅换填区", "—"),
        ],
    ),
    (
        "【三、旁压试验 Ménard（深度自地面起算）】",
        FILL_SECTION_GEOTECH,
        [],
    ),
    (
        "【四、本墩滑动与被动土压力】",
        FILL_SECTION_SLIDE,
        [
            ("hp_m", "", "被动区高度 hp (m)，留空则取 H", "m"),
            ("Kp", "", "Kp（留空按 Rankine 计算）", "—"),
            ("Fp_kN", "", "Fp (kN)，留空则自动计算", "kN"),
            ("glissement_include_C", 0, "滑动是否计 C'×A'/1.5（0/1）", "0/1"),
        ],
    ),
    (
        "【五、本墩沉降参数 Ménard】",
        FILL_SECTION_SETTLE,
        [
            ("Ec_kPa", "", "Ec 球形区；留空则由旁压 EM 计算", "kPa"),
            ("Ed_kPa", "", "Ed 偏应力区；留空则由旁压 EM 计算", "kPa"),
            ("alpha_settlement", 0.5, "流变系数 α", "—"),
            ("B0", 0.6, "参考宽度 B₀", "m"),
            ("lambda_c", 1.1, "形状系数 λc", "—"),
            ("lambda_d", 1.11, "形状系数 λd", "—"),
        ],
    ),
    (
        "【六、本墩配筋验算参数】",
        FILL_SECTION_REINF,
        [
            ("tau_ulim_kPa", 1260, "抗剪容许 τulim（对比 τu）", "kPa"),
            ("cover_top_cm", 9, "上缘保护层 c", "cm"),
            ("bar_dia_cm", 20, "主筋直径 φ（d=h−c−φ/2）", "cm"),
            ("gamma_beton_kN_m3", 25, "混凝土重度 γ_beton", "kN/m³"),
            ("ftj_MPa", 3.0, "混凝土 ftj（箍筋公式）", "MPa"),
            ("gamma_s", 1.15, "钢筋分项系数 γs", "—"),
            ("fe_MPa", 435, "钢筋 fe（箍筋公式）", "MPa"),
            ("bar_e_cm", 20, "主筋间距 e（计算书 HAxx(e=…)）", "cm"),
            ("rho_min_pct", 0.28, "最小配筋率（%）", "%"),
            ("midas_b_element_ids", "", "B-B 节点单元（Local V_q1/M_q1）", "—"),
            ("As_adopted_cm2_m", 40.2, "采用主筋面积（Expert）", "cm²/m"),
        ],
    ),
]

PIER_PARAM_KEYS = {p[0] for _, _, params in PIER_PARAM_SECTIONS for p in params}
STRING_PARAM_KEYS = {"soil_category", "ple_star_zone", "midas_element_ids", "midas_b_element_ids"}


def _style(cell, fill=None, font=None, align=None, border=True):
    if fill:
        cell.fill = fill
    if font:
        cell.font = font
    if align:
        cell.alignment = align
    if border:
        cell.border = BORDER_ALL


def _write_row(ws, row: int, values: List[Any], fills: Optional[List] = None, fonts: Optional[List] = None):
    for col, val in enumerate(values, 1):
        c = ws.cell(row=row, column=col, value=val)
        _style(
            c,
            fill=fills[col - 1] if fills else None,
            font=fonts[col - 1] if fonts else FONT_NORMAL,
            align=ALIGN_LEFT,
        )
    return row + 1


def _merge_section_title(ws, row: int, text: str, fill: PatternFill, ncol: int = 6) -> int:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncol)
    c = ws.cell(row=row, column=1, value=text)
    _style(c, fill=fill, font=FONT_SECTION, align=ALIGN_LEFT)
    return row + 1


def _set_col_widths(ws, widths: Dict[int, float]):
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w


def _write_midas_paste_block(ws, row: int, samples: List[List[Any]] = None) -> int:
    """【最下方】Midas 内力整表粘贴区。"""
    ncol = len(MIDAS_COLUMNS)
    title = "【七、MIDAS 内力粘贴】自 Midas 复制表头+数据整段粘贴到黄色区域（读取 load_mapping.txt 分类）"
    row = _merge_section_title(ws, row, title, FILL_SECTION_LOAD, ncol + 1)
    mcell = ws.cell(row=row - 1, column=ncol + 2, value=MIDAS_SECTION_KEY)
    mcell.font = Font(size=8, color="808080")
    cn_desc = ["单元", "荷载", "位置", "轴向", "剪力-y", "剪力-z", "扭矩", "弯矩-y", "弯矩-z"]
    units = ["—", "—", "—", "kN", "kN", "kN", "kN·m", "kN·m", "kN·m"]
    row = _write_row(ws, row, ["填写说明 →"] + cn_desc, [FILL_DESC] * (ncol + 1))
    row = _write_row(ws, row, ["单位"] + units, [FILL_DESC] * (ncol + 1))
    row = _write_row(
        ws, row, ["列名（勿改）"] + MIDAS_COLUMNS,
        [FILL_HEADER] + [FILL_HEADER] * ncol,
        [FONT_HEADER] + [FONT_HEADER] * ncol,
    )
    for sample in samples or []:
        row = _write_row(
            ws, row, [""] + list(sample),
            [FILL_DESC] + [FILL_INPUT] * ncol,
        )
    # 预留空行便于粘贴
    for _ in range(24):
        row = _write_row(
            ws, row, [""] + [""] * ncol,
            [FILL_DESC] + [FILL_INPUT] * ncol,
        )
    return row + 1


def _write_load_data_block(
    ws, row: int, marker: str, sec_title: str, sec_fill: PatternFill,
    cols: List[str], cn_desc: List[str], units: List[str], samples: List[Any],
) -> int:
    ncol = max(5, len(cols) + 1)
    row = _merge_section_title(ws, row, sec_title, sec_fill, ncol)
    mcell = ws.cell(row=row - 1, column=ncol + 1, value=marker)
    mcell.font = Font(size=8, color="808080")
    row = _write_row(ws, row, ["填写说明 →"] + cn_desc, [FILL_DESC] * (len(cols) + 1))
    row = _write_row(ws, row, ["单位"] + units, [FILL_DESC] * (len(cols) + 1))
    row = _write_row(
        ws, row, ["列名（勿改）"] + cols,
        [FILL_HEADER] + [FILL_HEADER] * len(cols),
        [FONT_HEADER] + [FONT_HEADER] * len(cols),
    )
    for sample in samples:
        row = _write_row(
            ws, row, [""] + list(sample),
            [FILL_DESC] + [FILL_INPUT] * len(cols),
        )
    return row + 1


PMT_BLOCK_DEF = (
    "Pressiometre_Menard",
    "【旁压试验数据】Pressiomètre Ménard",
    FILL_SECTION_GEOTECH,
    ["no", "depth_m", "lithology", "pf_MPa", "plm_MPa", "EM_MPa", "E_over_plm"],
    ["编号", "深度(地面起)", "岩性", "pf", "plm", "EM", "E/plm"],
    ["—", "m", "—", "MPa", "MPa", "MPa", "—"],
)


def _write_pier_params(ws, start_row: int, overrides: Dict[str, Any], pmt_samples: List[Any] = None) -> int:
    row = start_row
    pmt_marker, pmt_title, pmt_fill, pmt_cols, pmt_desc, pmt_units = PMT_BLOCK_DEF
    for idx, (sec_title, sec_fill, params) in enumerate(PIER_PARAM_SECTIONS):
        row = _merge_section_title(ws, row, sec_title, sec_fill, 4)
        if params:
            row = _write_row(
                ws, row,
                ["参数名（勿改）", "请输入数值 ▼", "中文说明", "单位"],
                [FILL_HEADER, FILL_HEADER, FILL_HEADER, FILL_HEADER],
                [FONT_HEADER, FONT_HEADER, FONT_HEADER, FONT_HEADER],
            )
            for key, default, desc, unit in params:
                val = overrides.get(key, default)
                row = _write_row(
                    ws, row,
                    [key, val, desc, unit],
                    [FILL_DESC, FILL_INPUT, FILL_DESC, FILL_DESC],
                )
        # 【三、旁压试验】标题下紧接着填表
        if sec_title.startswith("【三、旁压试验"):
            row = _write_load_data_block(
                ws, row, pmt_marker, pmt_title, pmt_fill,
                pmt_cols, pmt_desc, pmt_units, pmt_samples or [],
            )
        row += 1
    return row


# ---------- 算例荷载（墩柱1 / 墩柱2）----------
SAMPLE_ELU_1 = [
    (19871.74, 0.666, 0.0), (29185.94, 0.430, 0.008), (30042.70, 0.383, 0.0),
    (20456.16, 0.696, 0.010), (20697.38, 0.686, 0.044), (30474.07, 0.411, 0.0),
    (20169.90, 0.655, 0.012), (20456.16, 0.696, 0.010), (30042.70, 0.383, 0.0),
    (29723.92, 0.389, 0.031),
]
SAMPLE_ELS_1 = [
    (19902.78, 0.020, 0.0), (21643.09, 0.434, 0.008), (20963.83, 0.095, 0.0),
    (21742.03, 0.469, 0.007), (20680.25, 0.024, 0.033), (22719.79, 0.382, 0.0),
    (20169.90, 0.055, 0.008), (21742.03, 0.469, 0.007), (20963.83, 0.095, 0.0),
    (21981.08, 0.398, 0.031),
]
SAMPLE_HD_1 = [
    (3820.2,), (3815.8,), (3722.8,), (3908.7,), (3904.5,),
    (3812.2,), (3816.9,), (3908.7,), (3722.8,), (3727.2,),
]
SAMPLE_FERR_1 = [
    ("A-A", 1347.0, 896.0, 937.02),
    ("B-B", -334.1, 292.9, 537.20),
]

SAMPLE_ELU_2 = [
    (6887.37, 0.035, 0.0), (7138.14, 0.030, 0.132), (12376.53, 0.110, 0.0),
    (7828.33, 0.227, 0.039), (11680.64, 0.114, 0.123), (13561.59, 0.013, 0.0),
    (11170.34, 0.018, 0.085), (7828.33, 0.227, 0.039), (12376.53, 0.110, 0.0),
    (8423.16, 0.207, 0.170),
]
SAMPLE_ELS_2 = [
    (6832.49, 0.118, 0.0), (7138.14, 0.030, 0.088), (9154.11, 0.104, 0.0),
    (7711.28, 0.173, 0.030), (8551.79, 0.108, 0.123), (10155.90, 0.018, 0.0),
    (8360.12, 0.025, 0.075), (7711.28, 0.173, 0.030), (9154.11, 0.104, 0.0),
    (8238.03, 0.158, 0.128),
]
SAMPLE_HD_2 = [
    (22.2,), (121.9,), (124.3,), (161.7,), (120.9,),
    (16.0,), (121.7,), (161.7,), (124.3,), (158.3,),
]
SAMPLE_FERR_2 = [
    ("A-A", 502.6, 338.3, 490.33),
    ("B-B", 153.2, 196.7, -71.26),
]

SAMPLE_ELU_3 = [
    (6894.54, 0.035, 0.0), (11114.31, 0.016, 0.085), (7828.33, 0.227, 0.039),
    (12376.54, 0.110, 0.0), (8367.13, 0.205, 0.171), (13561.83, 0.013, 0.0),
    (7466.72, 0.029, 0.126), (12376.54, 0.110, 0.0), (7828.33, 0.227, 0.039),
    (12009.21, 0.111, 0.119),
]
SAMPLE_ELS_3 = [
    (6832.49, 0.118, 0.0), (8304.09, 0.022, 0.076), (7711.28, 0.173, 0.030),
    (9154.11, 0.104, 0.0), (8017.94, 0.146, 0.132), (10156.07, 0.018, 0.0),
    (7466.72, 0.029, 0.084), (9154.11, 0.104, 0.0), (7711.28, 0.173, 0.030),
    (9030.98, 0.089, 0.117),
]
SAMPLE_HD_3 = [
    (22.1,), (121.4,), (161.7,), (124.3,), (156.0,),
    (15.9,), (121.8,), (124.3,), (161.7,), (121.2,),
]
SAMPLE_FERR_3 = [
    ("A-A", 502.6, 338.3, 490.32),
    ("B-B", 153.7, 196.7, -71.26),
]

PIER1_PARAMS = {
    "H_cm": 130, "B_cm": 900, "L_cm": 920, "De_cm": 165, "Zw_cm": 0,
    "midas_element_ids": "5944",
    "gamma_kN_m3": 20, "kp": 1.02,
    "phi_deg": 25, "C_kPa": 75,
    "soil_category": "sable",
    "replacement_thickness_m": 1.30,
    "replacement_plm_MPa": 3.19,
    "ple_star_zone": "replacement",
    "hp_m": 1.30, "Kp": 2.464, "Fp_kN": 383.1,
    "alpha_settlement": 0.5,
    "B0": 0.6, "lambda_c": 1.1, "lambda_d": 1.11,
    "tau_ulim_kPa": 1260, "rho_min_pct": 0.28,
    "bar_e_cm": 20, "cover_top_cm": 9, "fe_MPa": 435, "As_adopted_cm2_m": 40.2,
}

PIER2_PARAMS = {
    "H_cm": 100, "B_cm": 500, "L_cm": 870, "De_cm": 115, "Zw_cm": 0,
    "gamma_kN_m3": 20, "kp": 1.03, "ple_star_kPa": 1871.08,
    "phi_deg": 25, "C_kPa": 75,
    "hp_m": 1.00, "Kp": 2.464, "Fp_kN": 214.4,
    "Ec_kPa": 25000, "Ed_kPa": 42290, "alpha_settlement": 0.5,
    "B0": 0.6, "lambda_c": 1.2, "lambda_d": 1.40,
    "tau_ulim_kPa": 1260, "rho_min_pct": 0.28,
    "bar_e_cm": 15, "cover_top_cm": 9, "fe_MPa": 435, "As_adopted_cm2_m": 32.7,
}

PIER3_PARAMS = {
    "H_cm": 100, "B_cm": 500, "L_cm": 870, "De_cm": 166, "Zw_cm": 0,
    "gamma_kN_m3": 20, "kp": 1.03, "ple_star_kPa": 2163.16,
    "phi_deg": 25, "C_kPa": 75,
    "hp_m": 1.00, "Kp": 2.464, "Fp_kN": 214.4,
    "Ec_kPa": 23567, "Ed_kPa": 43023, "alpha_settlement": 0.5,
    "B0": 0.6, "lambda_c": 1.2, "lambda_d": 1.40,
    "tau_ulim_kPa": 1260, "rho_min_pct": 0.28,
    "bar_e_cm": 15, "cover_top_cm": 9, "fe_MPa": 435, "As_adopted_cm2_m": 32.7,
}

# C0 桥台旁压试验样例（深度自地面 m）— (编号, depth, 岩性, pf, plm, EM, E/plm)
SAMPLE_PMT_C0 = [
    (1, 2.5, "Tufs", 1.15, 1.81, 20.9, 11.5),
    (2, 4.0, "Tufs", 0.58, 1.09, 13.3, 12.2),
    (3, 5.5, "Argile", 0.77, 1.42, 28.8, 20.3),
    (4, 7.0, "Argile", 1.00, 2.00, 21.2, 10.6),
    (5, 8.5, "Argile", 1.37, 1.78, 38.4, 21.6),
    (6, 10.0, "Argile", 1.81, 3.00, 43.9, 14.6),
    (7, 11.5, "Argile", 1.40, 2.36, 58.5, 24.8),
    (8, 13.0, "Argile", 1.06, 1.81, 19.0, 10.5),
    (9, 14.5, "Argile", 2.26, 3.66, 90.0, 24.6),
    (10, 16.0, "Argile", 2.08, 3.29, 52.3, 15.9),
    (11, 17.5, "Argile", 3.27, 5.00, 78.4, 15.7),
    (12, 19.0, "siltite", 3.91, 5.10, 199.0, 39.0),
    (13, 20.5, "siltite", 3.14, 5.11, 143.2, 28.0),
]


def write_pmt_sample_sheet(ws) -> None:
    """独立旁压试验样表（C0 桥台）— 可复制到各墩柱 sheet。"""
    ws.merge_cells("A1:G1")
    c = ws["A1"]
    c.value = "旁压试验样表 — C0 桥台（Pressiomètre Ménard）"
    _style(c, FILL_TITLE, FONT_TITLE, ALIGN_LEFT)
    ws.cell(row=2, column=1, value="说明：深度自地面起算 (m)；复制黄色区域到各墩【旁压试验数据】表格。")
    row = 4
    row = _write_row(
        ws, row,
        ["编号", "depth_m", "lithology", "pf_MPa", "plm_MPa", "EM_MPa", "E_over_plm"],
        [FILL_HEADER] * 7, [FONT_HEADER] * 7,
    )
    row = _write_row(
        ws, row,
        ["—", "深度 m", "岩性", "pf MPa", "plm MPa", "EM MPa", "E/plm"],
        [FILL_DESC] * 7,
    )
    for sample in SAMPLE_PMT_C0:
        row = _write_row(
            ws, row, list(sample),
            [FILL_DESC] + [FILL_INPUT] * 6,
        )
    _set_col_widths(ws, {1: 8, 2: 10, 3: 12, 4: 10, 5: 10, 6: 10, 7: 10})


def build_input_workbook(path: Path, pier_names: List[str]):
    wb = Workbook()
    ws0 = wb.active
    ws0.title = "使用说明"
    guide = [
        ("浅基础验算输入说明", ""),
        ("", ""),
        ("重要", "De = 基础埋深（地面至基底），非墩柱直径。q₀/ple*/Ec/Ed 可留空由程序自动计算。"),
        ("旁压试验", "墩柱1 土工参数下已附 C0 旁压表；q₀/ple*/Ec/Ed 留空则自动计算。"),
        ("Commun", "可选：全局默认参数（墩柱 sheet 中同名参数会覆盖）。"),
        ("墩柱1", "计算书 9.1.6 墩柱1 算例（10 组工况）。"),
        ("墩柱2", "计算书 2 号桥墩算例（10 组工况）。"),
        ("墩柱3", "计算书 3 号桥墩算例（10 组工况）。"),
        ("墩柱4～5", "空白模板。"),
        ("", ""),
        ("运行：python fondation_superficielle.py", ""),
        ("参数说明", "详见 sheet「参数取值说明」— 含取值来源与规范依据"),
    ]
    for i, (a, b) in enumerate(guide, 1):
        ws0.cell(row=i, column=1, value=a)
        ws0.cell(row=i, column=2, value=b)
    ws0.column_dimensions["A"].width = 18
    ws0.column_dimensions["B"].width = 72

    from param_reference import write_param_guide_sheet
    ws_guide = wb.create_sheet("参数取值说明")
    write_param_guide_sheet(ws_guide)

    ws = wb.create_sheet("Commun")
    ws.merge_cells("A1:D1")
    c = ws["A1"]
    c.value = "可选全局默认（墩柱 sheet 优先）— 一般留空即可"
    _style(c, FILL_TITLE, FONT_TITLE, ALIGN_LEFT)
    row = 3
    row = _write_row(ws, row, ["参数名", "默认值", "说明", "单位"], [FILL_HEADER] * 4, [FONT_HEADER] * 4)
    _set_col_widths(ws, {1: 22, 2: 16, 3: 42, 4: 10})

    load_block_defs = [
        ("ELU_mobilisation", "【ELU 竖向承载力】Mobilisation du sol", FILL_SECTION_LOAD,
         ["N_kN", "ex_m", "ey_m"], ["竖向力 N", "偏心 ex", "偏心 ey"], ["kN", "m", "m"]),
        ("ELS_mobilisation", "【ELS 竖向承载力】Mobilisation du sol", FILL_SECTION_LOAD,
         ["N_kN", "ex_m", "ey_m"], ["竖向力 N", "偏心 ex", "偏心 ey"], ["kN", "m", "m"]),
        ("ELU_glissement", "【ELU 滑动】Glissement", FILL_SECTION_SLIDE,
         ["Hd_kN"], ["水平力 Hd"], ["kN"]),
        ("Ferraillage", "【配筋】Ferraillage", FILL_SECTION_REINF,
         ["section", "Mmax_ELU_kNm_m", "Mmax_ELS_kNm_m", "Vd_ELU_kN"],
         ["截面", "Mmax ELU", "Mmax ELS", "Vd ELU"],
         ["—", "kN·m/m", "kN·m/m", "kN"]),
    ]

    # Midas 样例行（墩柱1）
    SAMPLE_MIDAS_1 = [
        (5944, "ELU包络(最大)", "I[2598]", 109.01, 102.42, 412.63, -60.88, 845.95, 90.68),
        (5944, "ELU包络(最小)", "I[2598]", 85.2, 50.1, 200.0, -30.0, 400.0, 40.0),
        (5944, "ELUA包络(最大)", "I[2598]", 95.0, 80.0, 350.0, -50.0, 700.0, 70.0),
        (5944, "ELS RARE包络(最大)", "I[2598]", 90.0, 60.0, 300.0, -40.0, 600.0, 60.0),
        (5944, "准永久组合", "I[2598]", 80.0, 40.0, 150.0, -20.0, 300.0, 30.0),
    ]

    pier_pmt_samples = {
        "墩柱1": SAMPLE_PMT_C0,
    }
    pier_load_samples = {
        "墩柱1": {
            "ELU_mobilisation": SAMPLE_ELU_1, "ELS_mobilisation": SAMPLE_ELS_1,
            "ELU_glissement": SAMPLE_HD_1, "Ferraillage": SAMPLE_FERR_1,
        },
        "墩柱2": {
            "ELU_mobilisation": SAMPLE_ELU_2, "ELS_mobilisation": SAMPLE_ELS_2,
            "ELU_glissement": SAMPLE_HD_2, "Ferraillage": SAMPLE_FERR_2,
        },
        "墩柱3": {
            "ELU_mobilisation": SAMPLE_ELU_3, "ELS_mobilisation": SAMPLE_ELS_3,
            "ELU_glissement": SAMPLE_HD_3, "Ferraillage": SAMPLE_FERR_3,
        },
    }
    pier_param_samples = {
        "墩柱1": PIER1_PARAMS, "墩柱2": PIER2_PARAMS, "墩柱3": PIER3_PARAMS,
    }

    for pier in pier_names:
        ws_p = wb.create_sheet(pier)
        ws_p.merge_cells("A1:F1")
        c = ws_p["A1"]
        c.value = f"{pier} — 参数 + 荷载（黄色单元格填写）"
        _style(c, FILL_TITLE, FONT_TITLE, ALIGN_LEFT)
        row = _write_pier_params(
            ws_p, 3, pier_param_samples.get(pier, {}),
            pier_pmt_samples.get(pier),
        )
        row += 1
        midas_samples = SAMPLE_MIDAS_1 if pier == "墩柱1" else []
        row = _write_midas_paste_block(ws_p, row, midas_samples)
        _set_col_widths(ws_p, {1: 14, 2: 18, 3: 14, 4: 14, 5: 14, 6: 14, 7: 14, 8: 14, 9: 14, 10: 14})

    wb.save(path)


def _mobil_row_out(r: Dict, regime: str) -> List[Any]:
    return [
        r["Cas"], r["N_kN"], r["ex_m"], r["ey_m"],
        r["q_ref1_kPa"], r["q_min_kPa"], r["q_max_kPa"],
        r["q_ref2_kPa"], r["q_ref_kPa"], r["gamma_q"], r["q0_kPa"],
        r["i_beta"], r["qu_kPa"], r["q_ref_lt_qu"],
    ]


MOBIL_HEADERS = [
    "", "N\n(kN)", "ex\n(m)", "ey\n(m)",
    "q'ref1(kPa)=\nN/(B-2ex)/(L-2ey)", "qmin\n(kPa)", "qmax\n(kPa)",
    "q'ref2(kPa)=\n(3qmax+qmin)/4", "q'ref\n(kPa)", "γq", "q'0\n(kPa)",
    "iδβ", "qu\n(kPa)", "q'ref<qu",
]


def build_output_workbook(path: Path, results: List[Dict[str, Any]]):
    wb = Workbook()
    wb.remove(wb.active)
    for res in results:
        name = res["pier_name"][:31]
        ws = wb.create_sheet(name)
        p = res["params"]
        kp = res.get("kp", "")
        ple = res.get("ple_star_kPa", "")
        q0 = p["q0_kPa"]
        q_u = p["q_u_kPa"]
        row = 1

        def title(text: str):
            nonlocal row
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=14)
            c = ws.cell(row=row, column=1, value=text)
            _style(c, FILL_OUTPUT_TITLE, FONT_OUTPUT_TITLE, ALIGN_LEFT)
            row += 1

        def section(text: str):
            nonlocal row
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=14)
            c = ws.cell(row=row, column=1, value=text)
            _style(c, FILL_OUTPUT_SECTION, FONT_OUTPUT_SECTION, ALIGN_LEFT)
            row += 1

        def table_header(headers: List[str]):
            nonlocal row
            for col, h in enumerate(headers, 1):
                c = ws.cell(row=row, column=col, value=h)
                _style(c, FILL_HEADER, FONT_HEADER, ALIGN_CENTER)
            row += 1

        def table_row(values: List[Any]):
            nonlocal row
            for col, v in enumerate(values, 1):
                c = ws.cell(row=row, column=col, value=v)
                _style(c, None, FONT_NORMAL, ALIGN_CENTER)
            row += 1

        title("9.1.6 Calcul de la fondation superficielle")
        section("1) Paramètres de calcul")
        geo_line = (
            f"Dimension de la fondation : H= {p['H_cm']} cm, B= {p['B_cm']} cm, "
            f"L= {p['L_cm']} cm, De= {p['De_cm']} cm (Profondeur d'encastrement), "
            f"Zw= {p.get('Zw_cm', 0)} cm ."
        )
        ga = res.get("geotech_auto", {})
        q0_src = ga.get("q0_source", "")
        q0_note = f" (auto γ×De)" if q0_src == "gamma_x_De" else ""
        if kp and ple:
            geotech_line = (
                f"Paramètres géotechniques : q'0= {q0}{q0_note} kPa, q'u = q0+ kp×ple*="
                f"{q0}+{kp}×{ple}= {q_u}kPa, φ = {p['phi_deg']}°, "
                f"C' = {p['C_kPa']} kPa, γ = {p['gamma_kN_m3']} kN/m3."
            )
        else:
            geotech_line = (
                f"Paramètres géotechniques : q'0= {q0}{q0_note} kPa, q'u= {q_u} kPa, "
                f"φ = {p['phi_deg']}°, C' = {p['C_kPa']} kPa, γ = {p['gamma_kN_m3']} kN/m3."
            )
        load_line = ""
        lm = res.get("load_meta", {})
        if lm.get("load_source") == "midas":
            load_line = (
                f"荷载 Midas：ELU {lm.get('elu_count', 0)} 组; "
                f"ELS fréquent(沉降) {lm.get('els_freq_count', 0)}; "
                f"ELS QP {lm.get('els_qp_count', 0)}; "
                f"ELS Rare(倾覆) {lm.get('els_rare_count', 0)} "
                f"(load_mapping.txt)"
            )
        menard_line = ""
        if ga:
            parts = []
            if ga.get("ple_source") == "menard":
                pm = ga.get("ple_meta") or {}
                zmode = pm.get("ple_star_zone", "mixed")
                if pm.get("replacement_used"):
                    parts.append(f"ple* Ménard + D2/D3 repl ({zmode})")
                else:
                    parts.append(f"ple* from Ménard (zone 1.5B)")
            if ga.get("Ec_source") == "menard":
                parts.append(f"Ec/Ed from Ménard EM layers")
            if parts:
                menard_line = "Géotech auto : " + "; ".join(parts) + "."
        phi1 = p.get("phi1_sand", "")
        b2de = p.get("B_over_2De", "")
        soil_cls = p.get("soil_class", "")
        mob0 = res["mobilisation_ELU"][0] if res.get("mobilisation_ELU") else {}
        i_formula = mob0.get("i_beta_formula", "")
        delta_deg = mob0.get("delta_deg", "")
        i_beta_val = mob0.get("i_beta", "")
        i_beta_line = (
            f"iδβ (Fasc.62 Annexe F.1, {soil_cls}, cas1 δ={delta_deg}°) : "
            f"{i_formula}; iδβ={i_beta_val}. "
            f"Φ1_sand(d配筋)={phi1} (B/(2De)={b2de})."
        )
        for line in [geo_line, geotech_line, i_beta_line] + (
            [load_line] if load_line else []
        ) + ([menard_line] if menard_line else []):
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=14)
            c = ws.cell(row=row, column=1, value=line)
            _style(c, None, FONT_NORMAL, ALIGN_LEFT)
            row += 1
        row += 1

        section("2) Vérification de l'assise")
        section("Vérification de la mobilisation du sol")
        table_header(["ELU"] + MOBIL_HEADERS[1:])
        for r in res["mobilisation_ELU"]:
            table_row(_mobil_row_out(r, "ELU"))
        row += 1
        if res.get("mobilisation_ELS"):
            section("ELS fréquent — q'ref 用于沉降（非 Rare）")
            table_header(["ELS"] + MOBIL_HEADERS[1:])
            for r in res["mobilisation_ELS"]:
                table_row(_mobil_row_out(r, "ELS"))
            row += 1

        section("Vérification au renversement")
        table_header(["ELU", "As", "As/A×100%", "≥10%"])
        for i, r in enumerate(res["renversement_ELU"], 1):
            table_row([i, r["As_m2"], r["As_over_A_pct"], r.get("renversement_ELU", "")])
        row += 1
        table_header(["ELS QP", "σmin", "σmin>0"])
        for i, r in enumerate(res["renversement_ELS_QP"], 1):
            table_row([i, r["sigma_min_kPa"], r.get("sigma_min_gt_0", "")])
        row += 1
        table_header(["ELS Rare", "As", "As/A×100%", "≥75%"])
        for i, r in enumerate(res["renversement_ELS_rare"], 1):
            table_row([i, r["As_m2"], r["As_over_A_pct"], r.get("renversement_ELS_rare", "")])
        row += 1

        section("Vérification au glissement")
        table_header([
            "ELU", "Hd\n(kN)", "hp\n(m)", "Kp", "Fp\n(kN)",
            "N*tan(φ)/1.2\n+C'A'/1.5", "Hd_lim", "Hd≤Hd_lim",
        ])
        for i, r in enumerate(res["glissement_ELU"], 1):
            comb = r["Ntan_phi_over_gamma_g1"] + r["Cprime_Aprime_over_gamma_g2"]
            table_row([
                i, r["Hd_kN"], r["hp_m"], r["Kp"], r["Fp_kN"],
                round(comb, 2), r["Hd_lim_kN"], r.get("Hd_le_Hd_lim", ""),
            ])
        row += 1

        section("Evaluation du tassement")
        table_header([
            "ELS", "B\n(m)", "L\n(m)", "q'ref\n（kPa）", "q0'\n（kPa）",
            "α", "Ec\n(kPa)", "Ed\n(kPa)", "B0\n(m)", "λc", "λd",
            "Sc\n(mm)", "Sd\n(mm)", "Sf\n(mm)",
        ])
        for i, r in enumerate(res["settlement_ELS"], 1):
            table_row([
                i, r["B_m"], r["L_m"], r["q_ref_kPa"], r["q0_kPa"],
                r["alpha"], r["Ec_kPa"], r["Ed_kPa"], r["B0"],
                r["lambda_c"], r["lambda_d"], r["Sc_mm"], r["Sd_mm"], r["Sf_mm"],
            ])
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=14)
        c = ws.cell(row=row, column=1, value=f"sfmax={res['sf_max_mm']}mm.")
        _style(c, None, FONT_NORMAL, ALIGN_LEFT)
        row += 2

        section("3) Ferraillage de la foundation")
        note = res.get("section_forces_note", "")
        if note:
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
            c = ws.cell(row=row, column=1, value=note)
            _style(c, None, FONT_NORMAL, ALIGN_LEFT)
            row += 2

        ferr = res.get("ferraillage", [])
        if ferr:
            hdr = row
            for col, h in enumerate(["section", "Mmax(kNm/m)", "", "As (cm2/m)"], 1):
                c = ws.cell(row=hdr, column=col, value=h)
                _style(c, FILL_HEADER, FONT_HEADER, ALIGN_CENTER)
            ws.merge_cells(start_row=hdr, start_column=2, end_row=hdr, end_column=3)
            row += 1
            for col, h in enumerate(["", "ELU", "ELS", ""], 1):
                c = ws.cell(row=row, column=col, value=h)
                _style(c, FILL_HEADER, FONT_HEADER, ALIGN_CENTER)
            row += 1
            for r in ferr:
                table_row([
                    r.get("section"),
                    r.get("Mmax_ELU_kNm_m"),
                    r.get("Mmax_ELS_kNm_m"),
                    r.get("As_cm2_m"),
                ])
            row += 1

            aa = next((x for x in ferr if x.get("section") == "A-A"), {})
            bb = next((x for x in ferr if x.get("section") == "B-B"), {})
            table_header(["section", "A-A", "B-B"])
            table_row(["Vd,max(kN)", "ELU", "ELU"])
            table_row(["", aa.get("Vd_max_ELU_kN"), bb.get("Vd_max_ELU_kN")])
            table_row(["τu(kPa)", aa.get("tau_u_kPa"), bb.get("tau_u_kPa")])
            table_row(["τulim(kPa)", aa.get("tau_ulim_kPa"), bb.get("tau_ulim_kPa")])
            table_row(["τu<τulim", aa.get("tau_ok"), bb.get("tau_ok")])
            row += 1

            for r in ferr:
                footer = r.get("footer_fr", "")
                if footer:
                    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
                    c = ws.cell(row=row, column=1, value=footer)
                    _style(c, None, FONT_NORMAL, ALIGN_LEFT)
                    row += 1
            row += 1
        _set_col_widths(ws, {i: 16 for i in range(1, 6)})
    from param_reference import write_output_calc_explanation
    write_output_calc_explanation(wb, results)
    wb.save(path)


def is_header_row(values: List[Any]) -> bool:
    headers = [str(v).strip() for v in values if v is not None and str(v).strip()]
    return any(h in KNOWN_COL_KEYS for h in headers)
