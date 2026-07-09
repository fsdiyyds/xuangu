# -*- coding: utf-8 -*-
"""
9.1.6 Calcul de la fondation superficielle（法国浅基础验算）

依据 SETRA / Fascicule 61 类公路桥梁浅基础常用验算条目实现，公式与您提供的
计算书样例对齐，便于对照核查。

本目录自包含：代码、input.xlsx、output.xlsx、理论说明 FONDATION_PRINCIPES.md

用法：
  cd expert/fondation_superficielle
  python fondation_superficielle.py
  python fondation_superficielle.py --create-template

输入：input.xlsx
  - Commun：几何、土工、材料、沉降参数
  - 墩柱1 … 墩柱5：各荷载工况（ELU/ELS）

输出：output.xlsx（每个墩柱一个 sheet）
"""

import argparse
import math
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

from excel_format import (
    build_input_workbook,
    build_output_workbook,
    is_header_row,
    LOAD_SECTION_KEYS,
    PIER_PARAM_KEYS,
    STRING_PARAM_KEYS,
    PRESSIOMETER_SECTION_KEY,
    ALL_SECTION_KEYS,
)
from menard_geo import normalize_pressiometer_rows, resolve_geotech_from_menard
from fascicule_i_beta import classify_soil, calc_i_delta_beta, SOIL_FRICTION
from midas_import import MIDAS_SECTION_KEY, resolve_pier_loads, _parse_element_filter
from section_forces import (
    GAMMA_B_ELUG,
    GAMMA_B_ELS,
    EXPERT_AS_LABEL,
    build_section_forces_from_midas,
    envelope_section_summary,
    to_ferraillage_output,
    ferraillage_footer_fr,
    format_bar_adopted_label,
    resolve_b1_b2_long_m,
    effective_depth_m,
)

SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT = SCRIPT_DIR / "input.xlsx"
DEFAULT_OUTPUT = SCRIPT_DIR / "output.xlsx"

PIER_SHEETS = ["墩柱1", "墩柱2", "墩柱3", "墩柱4", "墩柱5"]
GAMMA_Q_ELU = 2.0
GAMMA_Q_ELS = 3.0
GAMMA_G1 = 1.2   # 滑动摩擦项系数 γg1（Fascicule 62 B.3.4）
GAMMA_G2 = 1.5   # 滑动粘聚力项系数 γg2
B0_REF_M = 0.6   # 沉降参考宽度 B0（规范 0.60 m）


def _f(val, default=0.0) -> float:
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return default
    try:
        return float(val)
    except (TypeError, ValueError):
        return default


def _optional_float(val) -> Optional[float]:
    """空单元格 → None；有值 → float。"""
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return None
    if isinstance(val, str) and not str(val).strip():
        return None
    try:
        return float(val)
    except (TypeError, ValueError):
        return None


def read_commun(df: pd.DataFrame) -> Dict[str, float]:
    """Commun sheet：两列 参数名 / 数值"""
    params: Dict[str, float] = {}
    if df.shape[1] < 2:
        return params
    skip_keys = {"参数名（勿改）", "列名（勿改）", "填写说明 →", "单位", "参数名", "请输入数值 ▼"}
    for _, row in df.iterrows():
        key = str(row.iloc[0]).strip()
        if not key or key.startswith("#") or key.startswith("【") or key in skip_keys:
            continue
        if "说明" in key or "请输入" in key:
            continue
        params[key] = _f(row.iloc[1])
    return params


def read_load_table(df: pd.DataFrame) -> List[Dict[str, Any]]:
    if df.empty:
        return []
    df = df.dropna(how="all")
    rows: List[Dict[str, Any]] = []
    for _, row in df.iterrows():
        if pd.isna(row.iloc[0]):
            continue
        d = {str(c).strip(): row[c] for c in df.columns}
        rows.append(d)
    return rows


def meyerhof_compressed_area(
    ex_m: float, ey_m: float, B_m: float, L_m: float,
) -> Tuple[float, float, float, float, float]:
    """Meyerhof 有效受压区：B'=B-2ex, L'=L-2ey, As=B'×L'"""
    ex_m = max(ex_m, 0.0)
    ey_m = max(ey_m, 0.0)
    Bp = max(B_m - 2 * ex_m, 0.01)
    Lp = max(L_m - 2 * ey_m, 0.01)
    As = Bp * Lp
    A = B_m * L_m
    return As, A, As / A * 100.0, Bp, Lp


def rankine_kp(phi_deg: float) -> float:
    """Kp = tan²(45° + φ/2)"""
    return math.tan(math.radians(45.0 + phi_deg / 2.0)) ** 2


def calc_phi1_sand(B_m: float, De_m: float) -> float:
    """保留 B/(2De) 判定供配筋 d 折减等使用（非 iδβ）。"""
    if De_m <= 0 or B_m <= 0:
        return 1.0
    return 5.3 if B_m / (2.0 * De_m) > 2.5 else 1.0


def flex_depth_cm(
    section: str, H_cm: float, De_cm: float, bar_e_cm: float, B_cm: float = 0.0,
) -> float:
    """A-A：d=H−e；B-B：宽基础 B/(2De)>2.5 且埋深 De−H>30cm 时 d=(H−e)×H/De（De=埋深）"""
    d_aa = H_cm - bar_e_cm
    if section.strip().upper().startswith("B"):
        b2de = B_cm / (2.0 * De_cm) if B_cm > 0 and De_cm > 0 else 0.0
        if (
            b2de > 2.5
            and De_cm > H_cm
            and (De_cm - H_cm) > 30.0
        ):
            return max(d_aa * H_cm / De_cm, 1.0)
        return max(d_aa, 1.0)
    return max(d_aa, 1.0)


def flex_moment_knm(section: str, M_kNm_m: float, B_m: float, L_m: float) -> float:
    """B-B 纵向截面：|M|×1.2×L/B；A-A 横向：|M|"""
    m = abs(M_kNm_m)
    if section.strip().upper().startswith("B") and B_m > 0:
        return m * 1.2 * L_m / B_m
    return m


def mobilisation_row(
    case_no: int,
    regime: str,
    N_kN: float,
    ex_m: float,
    ey_m: float,
    B_m: float,
    L_m: float,
    De_m: float,
    q_u_kPa: float,
    q0_kPa: float,
    gamma_q: float,
    soil_class: str,
    phi_deg: float,
    C_kPa: float,
    Hd_kN: Optional[float] = None,
) -> Dict[str, Any]:
    """承载力 mobilisation du sol"""
    ex_m = max(ex_m, 0.0)
    ey_m = max(ey_m, 0.0)
    B_eff = max(B_m - 2 * ex_m, 0.01)
    L_eff = max(L_m - 2 * ey_m, 0.01)

    q_ref1 = N_kN / (B_eff * L_eff)
    q_mean = N_kN / (B_m * L_m)
    # 线性边缘应力：与计算书一致，仅沿 B 方向偏心 ex（ey 由 Meyerhof B'/L' 体现）
    q_min = q_mean * (1 - 6 * ex_m / B_m)
    q_max = q_mean * (1 + 6 * ex_m / B_m)
    q_ref2 = (3 * q_max + q_min) / 4.0
    q_ref = max(q_ref1, q_ref2)

    i_beta, delta_deg, phi_used, i_formula = calc_i_delta_beta(
        ex_m, B_m, De_m, soil_class, phi_deg, C_kPa, ey_m, L_m, Hd_kN, N_kN,
    )
    qu = (q_u_kPa - q0_kPa) * i_beta / gamma_q + q0_kPa
    ok = "OK" if q_ref < qu else "NON OK"

    return {
        "Cas": case_no,
        "Regime": regime,
        "N_kN": round(N_kN, 2),
        "ex_m": round(ex_m, 3),
        "ey_m": round(ey_m, 3),
        "q_ref1_kPa": round(q_ref1, 2),
        "q_min_kPa": round(q_min, 1),
        "q_max_kPa": round(q_max, 1),
        "q_ref2_kPa": round(q_ref2, 1),
        "q_ref_kPa": round(q_ref, 1),
        "gamma_q": gamma_q,
        "q0_kPa": q0_kPa,
        "i_beta": round(i_beta, 4),
        "delta_deg": delta_deg,
        "i_beta_formula": i_formula,
        "qu_kPa": round(qu, 1),
        "q_ref_lt_qu": ok,
    }


def renversement_check(
    As_m2: float,
    A_m2: float,
    min_pct: float,
    ok_key: str,
) -> Dict[str, Any]:
    """Fascicule 62 B.3.2/B.3.3：受压面积比 As/A"""
    pct = As_m2 / A_m2 * 100.0 if A_m2 > 0 else 0.0
    return {
        "As_m2": round(As_m2, 2),
        "As_over_A_pct": round(pct, 2),
        "As_min_pct": min_pct,
        ok_key: "OK" if pct >= min_pct else "NON OK",
    }


def renversement_els_qp(sigma_min_kPa: float) -> Dict[str, Any]:
    """ELS 准永久：基底全截面受压 σmin > 0"""
    ok = "OK" if sigma_min_kPa > 0 else "NON OK"
    return {"sigma_min_kPa": round(sigma_min_kPa, 1), "sigma_min_gt_0": ok}


def glissement_row(
    case_no: int,
    Hd_kN: float,
    N_kN: float,
    hp_m: float,
    Kp: float,
    L_m: float,
    phi_deg: float,
    C_kPa: float,
    gamma_kN_m3: float,
    A_prime_m2: float,
    Fp_override: Optional[float] = None,
    include_C: bool = False,
) -> Dict[str, Any]:
    """Fascicule 62 B.3.4 + Rankine 被动土压力 Fp"""
    phi = math.radians(phi_deg)
    if Fp_override is not None and Fp_override > 0:
        Fp = Fp_override
    else:
        Fp = 0.5 * Kp * gamma_kN_m3 * hp_m ** 2 * L_m
    tan_term = N_kN * math.tan(phi) / GAMMA_G1
    C_term = min(C_kPa, 75.0) * A_prime_m2 / GAMMA_G2 if include_C else 0.0
    Hd_lim = Fp + tan_term + C_term
    ok = "OK" if Hd_kN <= Hd_lim else "NON OK"
    return {
        "Cas": case_no,
        "Hd_kN": round(Hd_kN, 1),
        "hp_m": round(hp_m, 2),
        "Kp": round(Kp, 3),
        "Fp_kN": round(Fp, 1),
        "Ntan_phi_over_gamma_g1": round(tan_term, 2),
        "Cprime_Aprime_over_gamma_g2": round(C_term, 2),
        "Hd_lim_kN": round(Hd_lim, 1),
        "Hd_le_Hd_lim": ok,
    }


def settlement_row(
    case_no: int,
    B_m: float,
    L_m: float,
    q_ref_kPa: float,
    q0_kPa: float,
    alpha: float,
    Ec_kPa: float,
    Ed_kPa: float,
    B0: float,
    lambda_c: float,
    lambda_d: float,
) -> Dict[str, Any]:
    """Ménard 沉降：sc 用 Ec（球形区），sd 用 Ed（偏应力区）"""
    q_net = max(q_ref_kPa - q0_kPa, 0.0)
    if Ec_kPa <= 0 or Ed_kPa <= 0:
        Sc, Sd, Sf = 0.0, 0.0, 0.0
    else:
        ratio = lambda_d * B_m / B0 if B0 > 0 else 0.0
        Sc = alpha / (9.0 * Ec_kPa) * q_net * lambda_c * B_m * 1000.0
        Sd = 2.0 / (9.0 * Ed_kPa) * q_net * B0 * (ratio ** alpha) * 1000.0
        Sf = Sc + Sd
    return {
        "Cas": case_no,
        "B_m": B_m,
        "L_m": L_m,
        "q_ref_kPa": round(q_ref_kPa, 1),
        "q0_kPa": q0_kPa,
        "alpha": alpha,
        "Ec_kPa": Ec_kPa,
        "Ed_kPa": Ed_kPa,
        "B0": B0,
        "lambda_c": lambda_c,
        "lambda_d": lambda_d,
        "Sc_mm": round(Sc, 2),
        "Sd_mm": round(Sd, 2),
        "Sf_mm": round(Sf, 1),
    }


def reinforcement_section(
    section: str,
    Mmax_ELU: float,
    Mmax_ELS: float,
    Vd_ELU: float,
    H_cm: float,
    De_cm: float,
    B_cm: float,
    B_m: float,
    L_m: float,
    bar_e_cm: float,
    cover_top_cm: float,
    fe_MPa: float,
    tau_ulim: float,
    rho_min_pct: float,
    As_adopted_cm2_m: float,
) -> Dict[str, Any]:
    d_flex_cm = flex_depth_cm(section, H_cm, De_cm, bar_e_cm, B_cm)
    d_shear_m = max((H_cm - cover_top_cm) / 100.0, 0.01)
    tau_u = Vd_ELU / d_shear_m
    As_min = rho_min_pct / 100.0 * 100.0 * H_cm
    m_elu = flex_moment_knm(section, Mmax_ELU, B_m, L_m)
    As_from_M = (
        m_elu * 1e6 / (0.9 * d_flex_cm * 10.0 * fe_MPa) / 100.0
        if d_flex_cm > 0 and fe_MPa > 0 else 0.0
    )
    As_req = As_from_M
    ok_tau = "OK" if abs(tau_u) < tau_ulim else "NON OK"
    return {
        "section": section,
        "Mmax_ELU_kNm_m": round(Mmax_ELU, 1),
        "Mmax_ELS_kNm_m": round(Mmax_ELS, 1),
        "As_calc_cm2_m": round(As_req, 2),
        "As_min_cm2_m": round(As_min, 1),
        "As_adopted_cm2_m": As_adopted_cm2_m,
        "Vd_max_ELU_kN": round(Vd_ELU, 2),
        "tau_u_kPa": round(tau_u, 1),
        "tau_ulim_kPa": tau_ulim,
        "tau_ok": ok_tau,
    }


def ferraillage_from_input_table(
    pier_xl: Dict[str, pd.DataFrame],
    tau_ulim: float,
    H_cm: float,
    cover_top_cm: float,
    rho_min_pct: float,
    As_adopted_cm2_m: float,
    bar_dia_cm: float,
    bar_e_cm: float,
) -> List[Dict[str, Any]]:
    """手工 Ferraillage 表（无 Midas 截面力时）：弯矩/剪力取自 input，As 留 Expert。"""
    df = pier_xl.get("Ferraillage")
    if df is None or df.empty:
        return []
    as_min = round(rho_min_pct / 100.0 * 100.0 * H_cm, 1)
    bar_label = format_bar_adopted_label(bar_dia_cm, bar_e_cm)
    d_shear_m = max((H_cm - cover_top_cm) / 100.0, 0.01)
    rows: List[Dict[str, Any]] = []
    for _, row in df.iterrows():
        sec = str(row.get("section", "")).strip()
        if not sec:
            continue
        m_elu = _f(row.get("Mmax_ELU_kNm_m"))
        m_els = row.get("Mmax_ELS_kNm_m")
        m_els_v = _f(m_els) if m_els is not None and str(m_els).strip() != "" else ""
        vd = abs(_f(row.get("Vd_ELU_kN")))
        tau_u = round(vd / d_shear_m, 1) if d_shear_m > 0 else 0.0
        rows.append({
            "section": sec,
            "Mmax_ELU_kNm_m": round(m_elu, 1) if m_elu else m_elu,
            "Mmax_ELS_kNm_m": round(m_els_v, 1) if m_els_v != "" else "",
            "As_cm2_m": EXPERT_AS_LABEL,
            "As_min_cm2_m": as_min,
            "As_adopted_cm2_m": As_adopted_cm2_m,
            "Vd_max_ELU_kN": round(vd, 2),
            "tau_u_kPa": tau_u,
            "tau_ulim_kPa": tau_ulim,
            "tau_ok": "OK" if tau_u < tau_ulim else "NON OK",
            "footer_fr": ferraillage_footer_fr(
                sec, rho_min_pct, as_min, As_adopted_cm2_m, bar_label,
            ),
        })
    return rows


def _merge_params(commun: Dict[str, float], pier: Dict[str, float]) -> Dict[str, float]:
    """墩柱参数覆盖 Commun 可选默认值"""
    merged = dict(commun)
    for k, v in pier.items():
        if v is None or (isinstance(v, float) and math.isnan(v)):
            continue
        if isinstance(v, str) and not str(v).strip():
            continue
        merged[k] = v
    return merged


def calc_pier(
    pier_name: str,
    commun: Dict[str, float],
    pier_params: Dict[str, Any],
    pier_xl: Dict[str, pd.DataFrame],
) -> Dict[str, Any]:
    p = _merge_params(commun, pier_params)

    H_cm = _f(p.get("H_cm"), 130)
    B_cm = _f(p.get("B_cm"), 900)
    L_cm = _f(p.get("L_cm"), 920)
    De_cm = _f(p.get("De_cm"), 165)
    Zw_cm = _f(p.get("Zw_cm"), 0)

    gamma = _f(p.get("gamma_kN_m3"), 20)
    gamma_buoy_raw = _optional_float(p.get("gamma_buoy_kN_m3"))
    gamma_buoy = gamma_buoy_raw if gamma_buoy_raw is not None else max(gamma - 10.0, 0.0)

    B_m = B_cm / 100.0
    De_m = De_cm / 100.0
    Zw_m = Zw_cm / 100.0

    pmt_rows = normalize_pressiometer_rows(
        read_load_table(pier_xl.get(PRESSIOMETER_SECTION_KEY, pd.DataFrame()))
    )
    phi = _f(p.get("phi_deg"), 25)
    C = _f(p.get("C_kPa"), 75)
    repl_thick = _f(p.get("replacement_thickness_m"), 0.0)
    repl_plm = _optional_float(p.get("replacement_plm_MPa"))
    ple_zone = str(p.get("ple_star_zone") or "").strip()
    soil_cat_input = p.get("soil_category")
    soil_class = classify_soil(
        phi, C,
        str(soil_cat_input) if soil_cat_input is not None and str(soil_cat_input).strip() else None,
    )

    geo = resolve_geotech_from_menard(
        De_m, B_m, gamma, gamma_buoy, Zw_m, pmt_rows,
        _optional_float(p.get("q0_kPa")),
        _optional_float(p.get("ple_star_kPa")),
        _optional_float(p.get("Ec_kPa")),
        _optional_float(p.get("Ed_kPa")),
        repl_thick,
        repl_plm,
        ple_zone,
    )
    q0 = geo["q0_kPa"]
    ple_star = geo.get("ple_star_kPa")
    Ec = geo.get("Ec_kPa")
    Ed = geo.get("Ed_kPa")
    if ple_star is None:
        ple_star = _f(p.get("ple_star_kPa"), 0.0)
    if Ec is None:
        Ec = _f(p.get("Ec_kPa"), 0.0)
    if Ed is None:
        Ed = _f(p.get("Ed_kPa"), 0.0)

    kp = _f(p.get("kp"), 1.02)
    q_u_val = p.get("q_u_kPa")
    q_u = _f(q_u_val) if q_u_val and _f(q_u_val) > 0 else q0 + kp * ple_star

    hp_raw = _f(p.get("hp_m"), 0.0)
    hp = hp_raw if hp_raw > 0 else H_cm / 100.0
    Kp_input = _f(p.get("Kp"), 0.0)
    Kp = Kp_input if Kp_input > 0 else rankine_kp(phi)
    Fp_override = _f(p.get("Fp_kN"), 0.0)
    Fp_override = Fp_override if Fp_override > 0 else None
    gliss_include_C = _f(p.get("glissement_include_C"), 0) >= 1
    alpha_s = _f(p.get("alpha_settlement"), 0.5)
    B0 = _f(p.get("B0"), 0.6)
    lam_c = _f(p.get("lambda_c"), 1.1)
    lam_d = _f(p.get("lambda_d"), 1.11)

    tau_ulim = _f(p.get("tau_ulim_kPa"), 1260)
    rho_min = _f(p.get("rho_min_pct"), 0.28)
    As_adopted = _f(p.get("As_adopted_cm2_m"), 40.2)
    bar_e_cm = _f(p.get("bar_e_cm"), 15.0)
    cover_top_cm = _f(p.get("cover_top_cm"), 9.0)
    fe_MPa = _f(p.get("fe_MPa"), 435.0)

    L_m = L_cm / 100.0

    load_tables = resolve_pier_loads(pier_xl, p, B_m, L_m)
    load_meta = load_tables.get("load_meta", {})
    elu_loads = load_tables.get("ELU_mobilisation", [])
    els_freq_loads = load_tables.get("ELS_mobilisation_FREQ", [])
    if not els_freq_loads:
        els_freq_loads = load_tables.get("ELS_mobilisation", [])
    els_qp_loads = load_tables.get("ELS_mobilisation_QP", [])
    els_rare_loads = load_tables.get("ELS_mobilisation_RARE", [])
    elu_hd = load_tables.get("ELU_glissement", [])
    midas_elu_raw = load_tables.get("_midas_elu_raw", [])
    midas_els_freq_raw = load_tables.get("_midas_els_freq_raw", [])

    # --- Fig 8.2.3.1 纵桥向 A-A / B-B（N、My 除以横桥向 B → 单位宽度）---
    B1_m, B2_m = resolve_b1_b2_long_m(
        L_m,
        _optional_float(p.get("B_1_cm")),
        _optional_float(p.get("B_2_cm")),
        _optional_float(p.get("pier_a_cm")),
    )
    section_geometry_ok = B1_m is not None and B2_m is not None
    section_forces_note = ""
    if not section_geometry_ok:
        section_forces_note = (
            "未填写 pier_a_cm 或 B_1_cm/B_2_cm，跳过 A-A/B-B 截面力计算"
        )

    W_trans_m = B_m   # 横桥向宽度：单位宽度除数
    L_long_m = L_m    # 纵桥向长度：q 分布方向、B₁/B₂ 沿此向
    bar_dia_cm = _f(p.get("bar_dia_cm"), 20.0)
    gamma_beton = _f(p.get("gamma_beton_kN_m3"), 25.0)
    ftj_MPa = _f(p.get("ftj_MPa"), 3.0)
    gamma_s = _f(p.get("gamma_s"), 1.15)
    d_m = effective_depth_m(H_cm, cover_top_cm, bar_dia_cm)
    h_m = H_cm / 100.0

    cap_ids = _parse_element_filter(p.get("midas_element_ids"))
    b_ids = _parse_element_filter(p.get("midas_b_element_ids"))
    if not b_ids and cap_ids:
        b_ids = cap_ids  # 未填 B 节点单元时回退同一单元 Local 内力

    section_elu_cases: List[Dict] = []
    section_els_cases: List[Dict] = []
    if section_geometry_ok and midas_elu_raw:
        section_elu_cases = build_section_forces_from_midas(
            midas_elu_raw, cap_ids, b_ids,
            L_long_m, W_trans_m, B1_m, B2_m, h_m, d_m,
            gamma_beton, ftj_MPa, fe_MPa, gamma_s, tau_ulim,
            gamma_g=GAMMA_B_ELUG,
        )
    if section_geometry_ok and midas_els_freq_raw:
        section_els_cases = build_section_forces_from_midas(
            midas_els_freq_raw, cap_ids, b_ids,
            L_long_m, W_trans_m, B1_m, B2_m, h_m, d_m,
            gamma_beton, ftj_MPa, fe_MPa, gamma_s, tau_ulim,
            gamma_g=GAMMA_B_ELS,
        )
    section_envelope = envelope_section_summary(section_elu_cases)
    section_envelope_els = envelope_section_summary(section_els_cases)
    reinf = to_ferraillage_output(
        section_elu_cases, section_envelope, section_envelope_els,
        tau_ulim, H_cm, rho_min, As_adopted, bar_dia_cm, bar_e_cm,
    )
    if not reinf:
        reinf = ferraillage_from_input_table(
            pier_xl, tau_ulim, H_cm, cover_top_cm,
            rho_min, As_adopted, bar_dia_cm, bar_e_cm,
        )

    mobil_elu: List[Dict] = []
    for i, row in enumerate(elu_loads, 1):
        r = mobilisation_row(
            i, "ELU",
            _f(row.get("N_kN")),
            _f(row.get("ex_m")),
            _f(row.get("ey_m")),
            B_m, L_m, De_m, q_u, q0, GAMMA_Q_ELU,
            soil_class, phi, C,
        )
        if row.get("load_case"):
            r["load_case"] = row.get("load_case")
        mobil_elu.append(r)

    # ELS fréquentes：仅用于沉降 Evaluation du tassement（q'ref）
    mobil_els_freq: List[Dict] = []
    els_qref_map: Dict[int, float] = {}
    for i, row in enumerate(els_freq_loads, 1):
        r = mobilisation_row(
            i, "ELS",
            _f(row.get("N_kN")),
            _f(row.get("ex_m")),
            _f(row.get("ey_m")),
            B_m, L_m, De_m, q_u, q0, GAMMA_Q_ELS,
            soil_class, phi, C, None,
        )
        if row.get("load_case"):
            r["load_case"] = row.get("load_case")
        mobil_els_freq.append(r)
        els_qref_map[i] = r["q_ref_kPa"]

    # ELS QP：仅用于倾覆验算 σmin>0
    mobil_els_qp: List[Dict] = []
    els_qp_qmin_map: Dict[int, float] = {}
    for i, row in enumerate(els_qp_loads, 1):
        r = mobilisation_row(
            i, "ELS_QP",
            _f(row.get("N_kN")),
            _f(row.get("ex_m")),
            _f(row.get("ey_m")),
            B_m, L_m, De_m, q_u, q0, GAMMA_Q_ELS,
            soil_class, phi, C, None,
        )
        if row.get("load_case"):
            r["load_case"] = row.get("load_case")
        mobil_els_qp.append(r)
        els_qp_qmin_map[i] = r["q_min_kPa"]

    A_total = B_m * L_m

    renv_elu: List[Dict] = []
    for i, row in enumerate(elu_loads, 1):
        As, _, _, _, _ = meyerhof_compressed_area(
            _f(row.get("ex_m")), _f(row.get("ey_m")), B_m, L_m,
        )
        r = renversement_check(As, A_total, 10.0, "renversement_ELU")
        r["Cas"] = i
        renv_elu.append(r)

    renv_els_qp: List[Dict] = []
    for i in range(1, len(els_qp_loads) + 1):
        r = renversement_els_qp(els_qp_qmin_map.get(i, 0))
        r["Cas"] = i
        renv_els_qp.append(r)

    renv_els_rare: List[Dict] = []
    for i, row in enumerate(els_rare_loads, 1):
        As, _, _, _, _ = meyerhof_compressed_area(
            _f(row.get("ex_m")), _f(row.get("ey_m")), B_m, L_m,
        )
        r = renversement_check(As, A_total, 75.0, "renversement_ELS_rare")
        r["Cas"] = i
        renv_els_rare.append(r)

    hd_map = {i + 1: _f(row.get("Hd_kN")) for i, row in enumerate(elu_hd)}
    if len(hd_map) < len(elu_loads):
        for i, row in enumerate(elu_loads, 1):
            if i not in hd_map:
                hd_map[i] = _f(row.get("Hd_kN"))
    gliss: List[Dict] = []
    for i, row in enumerate(elu_loads, 1):
        Hd = hd_map.get(i, _f(row.get("Hd_kN")))
        As, _, _, _, _ = meyerhof_compressed_area(
            _f(row.get("ex_m")), _f(row.get("ey_m")), B_m, L_m,
        )
        gliss.append(
            glissement_row(
                i, Hd, _f(row.get("N_kN")), hp, Kp, L_m, phi, C, gamma, As,
                Fp_override, gliss_include_C,
            )
        )

    settlement: List[Dict] = []
    for i in range(1, len(els_freq_loads) + 1):
        settlement.append(
            settlement_row(
                i, B_m, L_m, els_qref_map.get(i, 0), q0,
                alpha_s, Ec, Ed, B0, lam_c, lam_d,
            )
        )
    sf_max = max((s["Sf_mm"] for s in settlement), default=0.0)

    reinf_out: List[Dict] = list(reinf)

    Fp_used = Fp_override if Fp_override is not None else 0.5 * Kp * gamma * hp ** 2 * L_m

    return {
        "pier_name": pier_name,
        "kp": kp,
        "ple_star_kPa": ple_star,
        "params": {
            "H_cm": H_cm, "B_cm": B_cm, "L_cm": L_cm, "De_cm": De_cm, "Zw_cm": Zw_cm,
            "q_u_kPa": round(q_u, 1), "q0_kPa": q0, "phi_deg": phi,
            "C_kPa": C, "gamma_kN_m3": gamma, "gamma_buoy_kN_m3": gamma_buoy,
            "B_over_2De": round(B_m / (2.0 * De_m), 3) if De_m > 0 else 0.0,
            "phi1_sand": calc_phi1_sand(B_m, De_m),
            "soil_class": soil_class,
            "replacement_thickness_m": repl_thick,
            "replacement_plm_MPa": repl_plm,
        },
        "geotech_auto": geo,
        "input_used": {
            "H_cm": H_cm, "B_cm": B_cm, "L_cm": L_cm, "De_cm": De_cm, "Zw_cm": Zw_cm,
            "q0_kPa": q0, "kp": kp, "ple_star_kPa": ple_star,
            "q_u_kPa": round(q_u, 1),
            "phi_deg": phi, "C_kPa": C, "gamma_kN_m3": gamma,
            "gamma_buoy_kN_m3": gamma_buoy,
            "hp_m": hp, "Kp": Kp, "Fp_kN": round(Fp_used, 1),
            "glissement_include_C": 1 if gliss_include_C else 0,
            "Ec_kPa": Ec, "Ed_kPa": Ed, "alpha_settlement": alpha_s,
            "B0": B0, "lambda_c": lam_c, "lambda_d": lam_d,
            "tau_ulim_kPa": tau_ulim, "rho_min_pct": rho_min,
            "bar_e_cm": bar_e_cm, "cover_top_cm": cover_top_cm,
            "fe_MPa": fe_MPa, "As_adopted_cm2_m": As_adopted,
        },
        "derived": {
            "q_u_kPa": round(q_u, 1),
            "phi1_sand": calc_phi1_sand(B_m, De_m),
            "B_over_2De": round(B_m / (2.0 * De_m), 3) if De_m > 0 else 0.0,
            "Kp_used": round(Kp, 4),
            "hp_used_m": round(hp, 2),
            "q0_calc_kPa": geo.get("q0_calc_kPa"),
            "ple_source": geo.get("ple_source"),
            "Ec_source": geo.get("Ec_source"),
            "Ed_source": geo.get("Ed_source"),
        },
        "mobilisation_ELU": mobil_elu,
        "mobilisation_ELS": mobil_els_freq,
        "mobilisation_ELS_QP": mobil_els_qp,
        "load_meta": load_meta,
        "renversement_ELU": renv_elu,
        "renversement_ELS_QP": renv_els_qp,
        "renversement_ELS_rare": renv_els_rare,
        "glissement_ELU": gliss,
        "settlement_ELS": settlement,
        "sf_max_mm": round(sf_max, 1),
        "section_forces_ELU": section_elu_cases,
        "section_forces_envelope": section_envelope,
        "section_forces_envelope_ELS": section_envelope_els,
        "section_forces_note": section_forces_note,
        "ferraillage": reinf_out,
    }


def create_input_template(path: Path):
    """生成分区配色、带中文说明的输入 Excel 模板。"""
    build_input_workbook(path, PIER_SHEETS)


def parse_pier_sheet(df: pd.DataFrame) -> Dict[str, Any]:
    """解析墩柱 sheet：顶部参数 + 荷载区块。"""
    params: Dict[str, float] = {}
    sections: Dict[str, pd.DataFrame] = {}
    current_key: Optional[str] = None
    header: Optional[List[str]] = None
    buffer: List[List[Any]] = []

    def flush():
        nonlocal buffer, header, current_key
        if current_key and header and buffer:
            sections[current_key] = pd.DataFrame(buffer, columns=header)
        buffer = []
        header = None

    skip_keys = {"参数名（勿改）", "列名（勿改）", "填写说明 →", "单位", "参数名", "请输入数值 ▼"}

    for _, row in df.iterrows():
        vals = [row.iloc[i] if i < len(row) else None for i in range(len(row))]
        key0 = str(vals[0]).strip() if vals[0] is not None and not pd.isna(vals[0]) else ""

        found_key = None
        for v in vals:
            s = str(v).strip() if v is not None and not pd.isna(v) else ""
            if s in ALL_SECTION_KEYS:
                found_key = s
                break
        if found_key:
            flush()
            current_key = found_key
            header = None
            continue
        if current_key and header is None:
            rest = [v for v in vals[1:]]
            if key0 == "列名（勿改）":
                header = [str(v).strip() for v in rest if v is not None and str(v).strip()]
                continue
            if is_header_row(rest):
                header = [str(v).strip() for v in rest if v is not None and str(v).strip()]
                continue
        if key0 in skip_keys:
            continue
        if key0.startswith("【"):
            flush()
            current_key = None
            continue
        if key0 in PIER_PARAM_KEYS:
            flush()
            current_key = None
            if key0 in STRING_PARAM_KEYS:
                if vals[1] is not None and not pd.isna(vals[1]):
                    params[key0] = str(vals[1]).strip()
            else:
                ov = _optional_float(vals[1])
                if ov is not None:
                    params[key0] = ov
            continue
        if current_key and header:
            if all(pd.isna(v) or str(v).strip() == "" for v in vals):
                continue
            buffer.append(vals[1:1 + len(header)])
            continue
    flush()
    return {"params": params, "sections": sections}


def run(input_path: Path, output_path: Path):
    if not input_path.is_file():
        create_input_template(input_path)
        print(f"已生成输入模板: {input_path}")
        print("请填写各墩柱数据后重新运行。")
        return

    commun_df = pd.read_excel(input_path, sheet_name="Commun", header=None)
    commun = read_commun(commun_df)

    results: List[Dict] = []
    for pier in PIER_SHEETS:
        try:
            raw = pd.read_excel(input_path, sheet_name=pier, header=None)
        except ValueError:
            print(f"  跳过：无 sheet {pier}")
            continue
        parsed = parse_pier_sheet(raw)
        pier_xl = parsed.get("sections", {})
        if not pier_xl and not parsed.get("params"):
            print(f"  跳过：{pier} 无有效数据")
            continue
        res = calc_pier(pier, commun, parsed.get("params", {}), pier_xl)
        lm = res.get("load_meta", {})
        if lm.get("load_source") == "midas":
            elu_n = lm.get("elu_count", 0)
            print(
                f"  荷载来源: Midas → ELU {elu_n} 组, "
                f"ELS(沉降) {lm.get('els_freq_count', 0)}, "
                f"ELS QP {lm.get('els_qp_count', 0)}, "
                f"ELS Rare(倾覆) {lm.get('els_rare_count', 0)}"
            )
            if elu_n == 0:
                print(
                    "  警告：Midas 数据已读取但未分到 ELU 工况，请检查 load_mapping.txt "
                    "（支持 ELU/ELS 或 ULS/SLS）及 midas_element_ids"
                )
            unmapped = lm.get("midas_unmapped") or []
            if unmapped:
                print(f"  警告：未映射荷载名（前5个）: {unmapped[:5]}")
        g = res.get("geotech_auto", {})
        if g.get("ple_source") == "missing" and not pier_xl.get(PRESSIOMETER_SECTION_KEY):
            print(f"  警告：{pier} 无旁压数据且未填 ple*")
        results.append(res)
        print(f"  完成: {pier}")

    if results:
        build_output_workbook(output_path, results)
    print(f"\n结果已保存: {output_path}")


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass

    parser = argparse.ArgumentParser(description="法国浅基础验算 9.1.6")
    parser.add_argument("--input", type=str, default=str(DEFAULT_INPUT))
    parser.add_argument("--output", type=str, default=str(DEFAULT_OUTPUT))
    parser.add_argument("--create-template", action="store_true", help="仅生成输入模板")
    args = parser.parse_args()

    input_path = Path(args.input).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()

    if args.create_template:
        create_input_template(input_path)
        print(f"模板: {input_path}")
        return

    print("=" * 55)
    print("浅基础验算 — 9.1.6 fondation superficielle")
    print("=" * 55)
    run(input_path, output_path)


if __name__ == "__main__":
    main()
