# -*- coding: utf-8 -*-
"""
Fig 8.2.3.1 — 纵桥向承台墩边截面 A-A / B-B（ELU Fondamental）

坐标约定（与 input 一致）：
  · L：纵桥向承台长度 → 基底应力分布方向，B₁/B₂/a 沿此向
  · B：横桥向承台宽度 → Midas 总内力 ÷ B 得单位宽度 (kN/m, kN·m/m)
  · 条带：1 m 横桥向宽 × 全纵桥向长

A-A：墩纵桥向一侧边线竖向面（靠 q_max）；B-B：另一侧（靠 q_min）。
"""

from __future__ import annotations

import math
from typing import Any, Dict, List, Optional, Tuple

GAMMA_B_ELUG = 1.5
GAMMA_B_ELS = 1.0
B0_SHEAR_M = 1.0
EXPERT_AS_LABEL = "需 Expert 计算"


def _f(v, default: float = 0.0) -> float:
    if v is None:
        return default
    try:
        x = float(v)
        return default if math.isnan(x) else x
    except (TypeError, ValueError):
        return default


def to_unit_width(
    N_total_kN: float,
    M_total_kNm: float,
    W_trans_m: float,
) -> Tuple[float, float, float]:
    """
    总内力 → 单位宽度（每米横桥向条带）。
    N' [kN/m], M' [kN·m/m], ex [m] = |M'|/N'
    """
    W = max(W_trans_m, 0.01)
    n_u = abs(N_total_kN) / W
    m_u = M_total_kNm / W
    ex = abs(m_u) / n_u if n_u > 1.0 else 0.0
    return n_u, m_u, ex


def base_stresses(
    N_u_kN_m: float,
    M_u_kNm_m: float,
    L_long_m: float,
    B1_m: float,
    B2_m: float,
) -> Dict[str, float]:
    """
    纵桥向基底线性应力 q_max, q_min, q_A, q_B（kPa）。
    N', M' 已为 kN/m、kN·m/m；L 为纵桥向全长。
    """
    L = max(L_long_m, 0.01)
    N = abs(N_u_kN_m)
    M = M_u_kNm_m

    q_moy = N / L
    delta_q = 6.0 * M / (L ** 2)
    q_max_lin = q_moy + delta_q
    q_min_lin = q_moy - delta_q

    if q_min_lin > 0:
        q_max = q_max_lin
        q_min = q_min_lin
        q_a = q_min + (q_max - q_min) * (L - B1_m) / L
        q_b = q_min + (q_max - q_min) * B2_m / L
        B0_eff = L
        mode = "full"
    else:
        denom = 6.0 * abs(M) + 3.0 * N * L
        B0_eff = (3.0 * N * L ** 2 / denom) if denom > 0 else L
        B0_eff = min(max(B0_eff, 0.01), L)
        q_max = 2.0 * N / B0_eff if B0_eff > 0 else 0.0
        q_min = 0.0
        x_comp_left = L - B0_eff
        q_a = q_max * max(B0_eff - B1_m, 0.0) / B0_eff
        b2_prime = max(B2_m - x_comp_left, 0.0)
        q_b = q_max * b2_prime / B0_eff if B0_eff > 0 else 0.0
        mode = "partial"

    return {
        "mode": mode,
        "q_moy_kPa": round(q_moy, 2),
        "q_max_kPa": round(q_max, 2),
        "q_min_kPa": round(q_min, 2),
        "q_A_kPa": round(q_a, 2),
        "q_B_kPa": round(q_b, 2),
        "B0_eff_m": round(B0_eff, 4),
    }


def footing_self_weight_kN_m(
    gamma_beton_kN_m3: float,
    h_m: float,
    strip_long_m: float,
    gamma_g: float = GAMMA_B_ELUG,
) -> float:
    """γ_beton·γ_g·h·B₁ → kN/m（1 m 横桥向条带，纵桥向长度 B₁）。"""
    return gamma_beton_kN_m3 * gamma_g * h_m * strip_long_m


def section_aa_forces(
    q_a: float,
    q_max: float,
    B1_m: float,
    h_m: float,
    gamma_beton_kN_m3: float,
    gamma_g: float = GAMMA_B_ELUG,
) -> Tuple[float, float]:
    soil_up = (q_a + q_max) / 2.0 * B1_m
    self_w = footing_self_weight_kN_m(gamma_beton_kN_m3, h_m, B1_m, gamma_g)
    v_aa = soil_up - self_w
    m_soil = (q_a + 2.0 * q_max) / 6.0 * B1_m ** 2
    m_self = self_w * B1_m / 2.0
    m_aa = m_soil - m_self
    return v_aa, m_aa


def section_bb_forces(
    q_min: float,
    q_b: float,
    B2_m: float,
    h_m: float,
    V_q1_kN_m: float,
    M_q1_kNm_m: float,
    gamma_beton_kN_m3: float,
    gamma_g: float = GAMMA_B_ELUG,
) -> Tuple[float, float]:
    v_qb = (q_min + q_b) / 2.0 * B2_m
    self_w = footing_self_weight_kN_m(gamma_beton_kN_m3, h_m, B2_m, gamma_g)
    v_bb = V_q1_kN_m - v_qb + self_w
    m_qb = (q_min + 2.0 * q_b) / 6.0 * B2_m ** 2
    m_self = self_w * B2_m ** 2 / 2.0
    m_bb = M_q1_kNm_m - m_qb + m_self
    return v_bb, m_bb


def tau_u_kpa(V_u_kN_m: float, d_m: float, b0_m: float = B0_SHEAR_M) -> float:
    """τ_u = V_u / (b₀·d)，V_u 为 kN/m，b₀=1 m → kPa"""
    if d_m <= 0 or b0_m <= 0:
        return 0.0
    return abs(V_u_kN_m) / (b0_m * d_m)


def stirrup_requirement(
    tau_u_kPa: float,
    ftj_MPa: float,
    fe_MPa: float,
    gamma_s: float = 1.15,
) -> Dict[str, float]:
    tau_mpa = tau_u_kPa / 1000.0
    rhs_mpa = max((tau_mpa - 0.3 * ftj_MPa) / 0.9, 0.0)
    at_st_req = rhs_mpa * gamma_s / fe_MPa * 100.0 if fe_MPa > 0 else 0.0
    ftj_kpa = ftj_MPa * 1000.0
    rhs_kpa = max((tau_u_kPa - 0.3 * ftj_kpa) / 0.9, 0.0)
    return {
        "tau_rhs_kPa": round(rhs_kpa, 2),
        "At_over_St_req_cm2_cm": round(at_st_req, 4),
    }


def compute_one_elu_case(
    load_case: str,
    N_total_kN: float,
    M_total_kNm: float,
    V_q1_total_kN: float,
    M_q1_total_kNm: float,
    L_long_m: float,
    W_trans_m: float,
    B1_m: float,
    B2_m: float,
    h_m: float,
    d_m: float,
    gamma_beton_kN_m3: float,
    ftj_MPa: float,
    fe_MPa: float,
    gamma_s: float,
    tau_ulim_kPa: float,
    gamma_g: float = GAMMA_B_ELUG,
) -> Dict[str, Any]:
    W = max(W_trans_m, 0.01)
    N_u, M_u, ex_m = to_unit_width(N_total_kN, M_total_kNm, W)
    V_q1_u = V_q1_total_kN / W
    M_q1_u = M_q1_total_kNm / W

    stresses = base_stresses(N_u, M_u, L_long_m, B1_m, B2_m)
    q_a = stresses["q_A_kPa"]
    q_max = stresses["q_max_kPa"]
    q_min = stresses["q_min_kPa"]
    q_b = stresses["q_B_kPa"]

    v_aa, m_aa = section_aa_forces(q_a, q_max, B1_m, h_m, gamma_beton_kN_m3, gamma_g)
    v_bb, m_bb = section_bb_forces(
        q_min, q_b, B2_m, h_m, V_q1_u, M_q1_u, gamma_beton_kN_m3, gamma_g,
    )

    tau_aa = tau_u_kpa(v_aa, d_m)
    tau_bb = tau_u_kpa(v_bb, d_m)
    stir_aa = stirrup_requirement(tau_aa, ftj_MPa, fe_MPa, gamma_s)
    stir_bb = stirrup_requirement(tau_bb, ftj_MPa, fe_MPa, gamma_s)

    return {
        "load_case": load_case,
        "N_total_kN": round(N_total_kN, 2),
        "M_total_kNm": round(M_total_kNm, 2),
        "N_Ed_kN_m": round(N_u, 2),
        "M_Ed_kNm_m": round(M_u, 2),
        "ex_m": round(ex_m, 4),
        "V_q1_kN_m": round(V_q1_u, 2),
        "M_q1_kNm_m": round(M_q1_u, 2),
        **stresses,
        "V_AA_kN_m": round(v_aa, 2),
        "M_AA_kNm_m": round(m_aa, 2),
        "V_BB_kN_m": round(v_bb, 2),
        "M_BB_kNm_m": round(m_bb, 2),
        "tau_u_A_kPa": round(tau_aa, 1),
        "tau_u_B_kPa": round(tau_bb, 1),
        "tau_ulim_kPa": tau_ulim_kPa,
        "tau_A_ok": "OK" if tau_aa < tau_ulim_kPa else "NON OK",
        "tau_B_ok": "OK" if tau_bb < tau_ulim_kPa else "NON OK",
        "At_St_req_A": stir_aa["At_over_St_req_cm2_cm"],
        "At_St_req_B": stir_bb["At_over_St_req_cm2_cm"],
        "L_long_m": L_long_m,
        "W_trans_m": W_trans_m,
        "B1_m": B1_m,
        "B2_m": B2_m,
        "d_m": d_m,
    }


def resolve_b1_b2_long_m(
    L_long_m: float,
    B1_cm: Optional[float],
    B2_cm: Optional[float],
    pier_a_cm: Optional[float],
) -> Tuple[Optional[float], Optional[float]]:
    """
    B₁、B₂：纵桥向墩边至承台外缘 (m)。
    必须填写 pier_a_cm，或同时填写 B_1_cm 与 B_2_cm；否则返回 (None, None)。
    """
    has_a = pier_a_cm is not None and pier_a_cm > 0
    has_b1 = B1_cm is not None and B1_cm > 0
    has_b2 = B2_cm is not None and B2_cm > 0

    if has_b1 and has_b2:
        return B1_cm / 100.0, B2_cm / 100.0

    if has_a:
        a_m = pier_a_cm / 100.0
        rest = max(L_long_m - a_m, 0.02)
        b1 = B1_cm / 100.0 if has_b1 else rest / 2.0
        b2 = B2_cm / 100.0 if has_b2 else rest / 2.0
        return b1, b2

    return None, None


def effective_depth_m(h_cm: float, cover_top_cm: float, bar_dia_cm: float) -> float:
    d_cm = h_cm - cover_top_cm - bar_dia_cm / 2.0
    return max(d_cm / 100.0, 0.01)


def build_section_forces_from_midas(
    midas_rows: List[Dict[str, Any]],
    cap_element_ids: Optional[set],
    b_element_ids: Optional[set],
    L_long_m: float,
    W_trans_m: float,
    B1_m: float,
    B2_m: float,
    h_m: float,
    d_m: float,
    gamma_beton_kN_m3: float,
    ftj_MPa: float,
    fe_MPa: float,
    gamma_s: float,
    tau_ulim_kPa: float,
    gamma_g: float = GAMMA_B_ELUG,
) -> List[Dict[str, Any]]:
    from collections import defaultdict

    global_by_lc: Dict[str, Dict[str, float]] = defaultdict(lambda: {"N": 0.0, "M": 0.0, "n": 0})
    local_by_lc: Dict[str, Dict[str, float]] = {}

    for row in midas_rows:
        lc = str(row.get("load_case", "")).strip()
        if not lc:
            continue
        elem = str(row.get("element", "")).strip()
        in_cap = (not cap_element_ids) or (elem in cap_element_ids)

        if in_cap:
            global_by_lc[lc]["N"] += _f(row.get("N_kN"))
            global_by_lc[lc]["M"] += _f(row.get("moment_y_kNm"))
            global_by_lc[lc]["n"] += 1

        if b_element_ids and elem in b_element_ids:
            local_by_lc[lc] = {
                "V_q1": _f(row.get("shear_y_kN")),
                "M_q1": _f(row.get("moment_y_kNm")),
            }

    results: List[Dict[str, Any]] = []
    for lc, g in sorted(global_by_lc.items()):
        if g["n"] == 0:
            continue
        loc = local_by_lc.get(lc, {})
        results.append(
            compute_one_elu_case(
                lc, g["N"], g["M"],
                loc.get("V_q1", 0.0), loc.get("M_q1", 0.0),
                L_long_m, W_trans_m, B1_m, B2_m, h_m, d_m,
                gamma_beton_kN_m3, ftj_MPa, fe_MPa, gamma_s, tau_ulim_kPa,
                gamma_g,
            )
        )
    return results


def envelope_section_summary(cases: List[Dict[str, Any]]) -> Dict[str, Any]:
    """纯弯控制：各截面取 |M| 最大工况的带符号弯矩；剪力取 |V| 最大工况的 V 与对应 τu。"""
    if not cases:
        return {}

    def _ctrl_m(key: str) -> float:
        c = max(cases, key=lambda x: abs(x.get(key, 0)))
        return round(c[key], 1)

    def _ctrl_v_tau(v_key: str, tau_key: str) -> Tuple[float, float]:
        c = max(cases, key=lambda x: abs(x.get(v_key, 0)))
        return round(abs(c[v_key]), 2), round(c[tau_key], 1)

    v_aa, tau_a = _ctrl_v_tau("V_AA_kN_m", "tau_u_A_kPa")
    v_bb, tau_b = _ctrl_v_tau("V_BB_kN_m", "tau_u_B_kPa")
    return {
        "M_AA_ctrl_kNm_m": _ctrl_m("M_AA_kNm_m"),
        "M_BB_ctrl_kNm_m": _ctrl_m("M_BB_kNm_m"),
        "V_AA_max_kN_m": v_aa,
        "V_BB_max_kN_m": v_bb,
        "tau_u_A_kPa": tau_a,
        "tau_u_B_kPa": tau_b,
    }


def format_bar_adopted_label(bar_dia_cm: float, bar_e_cm: float) -> str:
    """输出计算书采用钢筋描述，如 HA32(e=20cm)。"""
    if bar_dia_cm >= 8:
        dia_mm = int(round(bar_dia_cm))
    else:
        dia_mm = int(round(bar_dia_cm * 10))
    return f"HA{dia_mm}(e={int(bar_e_cm)}cm)"


def ferraillage_footer_fr(
    section: str,
    rho_min_pct: float,
    as_min_cm2_m: float,
    as_adopted_cm2_m: float,
    bar_label: str,
) -> str:
    return (
        f"La section {section} maximale d'armature est de A = {EXPERT_AS_LABEL}, "
        f"le taux minimal de ferraillage exigée > {rho_min_pct}% = {as_min_cm2_m} cm2/m, "
        f"la section d'armatures adoptée est de A = {bar_label} = {as_adopted_cm2_m} cm2/m."
    )


def to_ferraillage_output(
    elu_cases: List[Dict[str, Any]],
    elu_envelope: Dict[str, Any],
    els_envelope: Dict[str, Any],
    tau_ulim_kPa: float,
    H_cm: float,
    rho_min_pct: float,
    As_adopted_cm2_m: float,
    bar_dia_cm: float,
    bar_e_cm: float,
) -> List[Dict[str, Any]]:
    if not elu_envelope:
        return []

    as_min = round(rho_min_pct / 100.0 * 100.0 * H_cm, 1)
    bar_label = format_bar_adopted_label(bar_dia_cm, bar_e_cm)
    specs = [
        ("A-A", "M_AA_ctrl_kNm_m", "V_AA_max_kN_m", "tau_u_A_kPa", "At_St_req_A"),
        ("B-B", "M_BB_ctrl_kNm_m", "V_BB_max_kN_m", "tau_u_B_kPa", "At_St_req_B"),
    ]
    rows: List[Dict[str, Any]] = []
    for sec, m_key, v_key, tau_key, at_key in specs:
        m_els = els_envelope.get(m_key, "") if els_envelope else ""
        tau_u = elu_envelope.get(tau_key, 0)
        rows.append({
            "section": sec,
            "Mmax_ELU_kNm_m": elu_envelope.get(m_key, ""),
            "Mmax_ELS_kNm_m": m_els,
            "As_cm2_m": EXPERT_AS_LABEL,
            "As_min_cm2_m": as_min,
            "As_adopted_cm2_m": As_adopted_cm2_m,
            "Vd_max_ELU_kN": elu_envelope.get(v_key, 0),
            "tau_u_kPa": tau_u,
            "tau_ulim_kPa": tau_ulim_kPa,
            "tau_ok": "OK" if tau_u < tau_ulim_kPa else "NON OK",
            "At_St_req_cm2_cm": max((c.get(at_key, 0) for c in elu_cases), default=0),
            "footer_fr": ferraillage_footer_fr(
                sec, rho_min_pct, as_min, As_adopted_cm2_m, bar_label,
            ),
        })
    return rows
