# -*- coding: utf-8 -*-
"""
Fascicule 62 / DTU 13-12 — 旁压试验解释与土工参数自动计算

- q'₀ = σ'v0(De) ：基底有效自重应力
- pl* = plm − σ'v0(z) ；ple* = 基底下 0～1.5B 区内 pl* 算术平均（含 1.5×pl*min 上限）
- Ec = E1 ：基底下 0～B/2 区 EM 调和平均
- Ed ：按 DTU/Fasc.62 五层调和平均公式（式 2.44）
"""

from __future__ import annotations

import math
from typing import Any, Dict, List, Optional, Tuple

# Ed 分层：深度以 B/2 为单位（基底以下相对深度 z/B）
ED_LAYER_BOUNDS = [
    (0.0, 1.0),    # E1 → Ec
    (1.0, 2.0),    # E2
    (2.0, 3.5),    # E3.5
    (3.5, 6.8),    # E6.8
    (6.8, 9.16),   # E9.16
]
ED_LAYER_WEIGHTS = [1.0, 0.85, 1.0, 2.5, 2.5]  # 4/Ed = Σ 1/(w_i * Ei)


def effective_stress_vertical_kpa(
    depth_m: float,
    gamma_kN_m3: float,
    gamma_buoy_kN_m3: float,
    zw_m: float,
) -> float:
    """自地面起算深度 z 处竖向有效应力 σ'v0 (kPa)。"""
    if depth_m <= 0:
        return 0.0
    if zw_m <= 0 or depth_m <= zw_m:
        return gamma_kN_m3 * depth_m
    return gamma_kN_m3 * zw_m + gamma_buoy_kN_m3 * (depth_m - zw_m)


def calc_q0_kpa(
    embedment_m: float,
    gamma_kN_m3: float,
    gamma_buoy_kN_m3: float,
    zw_m: float,
) -> float:
    """q'₀ = σ'v0(De)，De 为地面至基底埋深 (m)。"""
    return effective_stress_vertical_kpa(embedment_m, gamma_kN_m3, gamma_buoy_kN_m3, zw_m)


def _interp_at_depth(
    tests: List[Dict[str, float]], depth_m: float, key: str,
) -> Optional[float]:
    """按深度线性插值 plm_MPa 或 EM_MPa。"""
    pts = sorted(
        [(t["depth_m"], t[key]) for t in tests if t.get(key) is not None],
        key=lambda x: x[0],
    )
    if not pts:
        return None
    if depth_m <= pts[0][0]:
        return pts[0][1]
    if depth_m >= pts[-1][0]:
        return pts[-1][1]
    for i in range(len(pts) - 1):
        z0, v0 = pts[i]
        z1, v1 = pts[i + 1]
        if z0 <= depth_m <= z1:
            if z1 == z0:
                return v0
            t = (depth_m - z0) / (z1 - z0)
            return v0 + t * (v1 - v0)
    return pts[-1][1]


def pl_star_kpa_at_depth(
    depth_m: float,
    plm_MPa: float,
    gamma_kN_m3: float,
    gamma_buoy_kN_m3: float,
    zw_m: float,
) -> float:
    plm_kpa = plm_MPa * 1000.0
    sigma_v0 = effective_stress_vertical_kpa(depth_m, gamma_kN_m3, gamma_buoy_kN_m3, zw_m)
    return max(plm_kpa - sigma_v0, 0.0)


def _values_in_zone(
    tests: List[Dict[str, float]],
    z_min: float,
    z_max: float,
    value_key: str,
    fallback_key: Optional[str] = None,
) -> List[float]:
    """取试验点或边界插值点，用于层内平均。"""
    vals: List[float] = []
    for t in tests:
        z = t["depth_m"]
        if z_min <= z <= z_max:
            v = t.get(value_key)
            if v is not None and v > 0:
                vals.append(v)
    if vals:
        return vals
    if fallback_key:
        for z_bound in (z_min, z_max, (z_min + z_max) / 2.0):
            v = _interp_at_depth(tests, z_bound, fallback_key)
            if v is not None and v > 0:
                vals.append(v)
    return vals


def harmonic_mean(values: List[float]) -> Optional[float]:
    pos = [v for v in values if v > 0]
    if not pos:
        return None
    return len(pos) / sum(1.0 / v for v in pos)


def arithmetic_mean(values: List[float]) -> Optional[float]:
    if not values:
        return None
    return sum(values) / len(values)


def calc_ple_star_kpa(
    tests: List[Dict[str, float]],
    embedment_m: float,
    B_m: float,
    gamma_kN_m3: float,
    gamma_buoy_kN_m3: float,
    zw_m: float,
    replacement_thickness_m: float = 0.0,
    replacement_plm_MPa: Optional[float] = None,
    ple_star_zone: str = "mixed",
) -> Tuple[Optional[float], Dict[str, Any]]:
    """
    ple*：基底 De 至 De+1.5B 内 pl* 算术平均（DTU 13-12 / Fasc. 62）。

    换填层（D2/D3 等）：在基底以下 replacement_thickness_m 深度内，
    用换填材料 plm（replacement_plm_MPa）计算 pl*，而非原位粘土旁压值。

    ple_star_zone：
      - mixed：换填区用换填 plm，其下 1.5B 内仍计原位旁压（默认混合）。
      - replacement：仅换填厚度内（限 1.5B）用换填 plm 平均，不计下方原位土
        （承台底换填 D2/D3、基础直接坐落在换填层上时常用手算做法）。
    """
    zone_top = embedment_m
    zone_bottom = embedment_m + 1.5 * B_m
    zone_mode = str(ple_star_zone or "mixed").strip().lower()
    replacement_only = zone_mode.startswith("repl")
    meta: Dict[str, Any] = {
        "zone_m": [zone_top, zone_bottom],
        "n_points": 0,
        "replacement_used": False,
        "ple_star_zone": "replacement" if replacement_only else "mixed",
    }
    if B_m <= 0:
        return None, meta

    repl_thick = max(replacement_thickness_m, 0.0)
    repl_plm = replacement_plm_MPa
    repl_bottom = embedment_m + repl_thick if repl_thick > 0 else embedment_m
    if replacement_only and repl_thick > 0:
        zone_bottom_eff = min(embedment_m + repl_thick, zone_bottom)
    else:
        zone_bottom_eff = zone_bottom

    pl_stars: List[float] = []

    # 换填区：在承载力影响深度内用 D2/D3 砂砾石 plm
    if repl_plm is not None and repl_plm > 0 and repl_thick > 0:
        meta["replacement_used"] = True
        meta["replacement_zone_m"] = [embedment_m, min(repl_bottom, zone_bottom_eff)]
        # 换填区顶、中、底（限在 zone 内）取代表点
        repl_z_samples = {embedment_m}
        if repl_bottom > embedment_m:
            repl_z_samples.add((embedment_m + repl_bottom) / 2.0)
            repl_z_samples.add(min(repl_bottom, zone_bottom_eff))
        for z in sorted(repl_z_samples):
            if zone_top <= z <= zone_bottom_eff:
                pl_stars.append(
                    pl_star_kpa_at_depth(z, repl_plm, gamma_kN_m3, gamma_buoy_kN_m3, zw_m)
                )

    # 换填以下至 1.5B：原位旁压试验（replacement 模式跳过）
    nat_z_min = max(repl_bottom, zone_top)
    if not replacement_only and tests:
        for t in tests:
            z = t["depth_m"]
            if nat_z_min <= z <= zone_bottom and t.get("plm_MPa") is not None:
                pl_stars.append(
                    pl_star_kpa_at_depth(z, t["plm_MPa"], gamma_kN_m3, gamma_buoy_kN_m3, zw_m)
                )

    if not pl_stars and tests and not replacement_only:
        z_mid = (zone_top + zone_bottom) / 2.0
        if z_mid >= nat_z_min:
            plm = _interp_at_depth(tests, z_mid, "plm_MPa")
            if plm is not None:
                pl_stars.append(
                    pl_star_kpa_at_depth(z_mid, plm, gamma_kN_m3, gamma_buoy_kN_m3, zw_m)
                )

    if not pl_stars:
        return None, meta

    pl_min = min(pl_stars)
    cap = 1.5 * pl_min
    capped = [min(p, cap) for p in pl_stars]
    meta["n_points"] = len(pl_stars)
    meta["pl_star_min_kPa"] = round(pl_min, 1)
    return round(arithmetic_mean(capped), 2), meta


def _layer_em_kpa(
    tests: List[Dict[str, float]],
    embedment_m: float,
    B_m: float,
    z_lo_frac: float,
    z_hi_frac: float,
) -> Optional[float]:
    """单层 EM 调和平均（EM 输入 MPa，输出 kPa）。"""
    half_b = B_m / 2.0
    z_min = embedment_m + z_lo_frac * half_b
    z_max = embedment_m + z_hi_frac * half_b
    em_mpa_vals = _values_in_zone(tests, z_min, z_max, "EM_MPa", "EM_MPa")
    hm = harmonic_mean(em_mpa_vals)
    return hm * 1000.0 if hm is not None else None


def calc_ec_ed_kpa(
    tests: List[Dict[str, float]],
    embedment_m: float,
    B_m: float,
) -> Tuple[Optional[float], Optional[float], Dict[str, Any]]:
    """
    Ec = E1（0～B/2  beneath base）
    Ed = 4 / (1/E1 + 1/(0.85 E2) + 1/E3.5 + 1/(2.5 E6.8) + 1/(2.5 E9.16))
    """
    meta: Dict[str, Any] = {"layers_kPa": {}}
    if not tests or B_m <= 0:
        return None, None, meta
    layer_em: List[Optional[float]] = []
    names = ["E1", "E2", "E3.5", "E6.8", "E9.16"]
    for (lo, hi), name in zip(ED_LAYER_BOUNDS, names):
        em_kpa = _layer_em_kpa(tests, embedment_m, B_m, lo, hi)
        layer_em.append(em_kpa)
        if em_kpa is not None:
            meta["layers_kPa"][name] = round(em_kpa, 1)
    e1 = layer_em[0]
    if e1 is None:
        return None, None, meta
    ec = e1
    inv_sum = 0.0
    n_used = 0
    for em, w in zip(layer_em, ED_LAYER_WEIGHTS):
        if em is None or em <= 0:
            continue
        inv_sum += 1.0 / (w * em)
        n_used += 1
    ed = (4.0 / inv_sum) if inv_sum > 0 and n_used >= 2 else e1
    return round(ec, 1), round(ed, 1), meta


def normalize_pressiometer_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, float]]:
    """解析 Excel 旁压试验表。"""
    out: List[Dict[str, float]] = []
    for row in rows:
        depth = row.get("depth_m")
        if depth is None:
            continue
        try:
            d = float(depth)
        except (TypeError, ValueError):
            continue
        if math.isnan(d):
            continue
        item: Dict[str, float] = {"depth_m": d}
        for key in ("pf_MPa", "plm_MPa", "EM_MPa", "E_over_plm"):
            v = row.get(key)
            if v is None or (isinstance(v, float) and math.isnan(v)):
                continue
            try:
                fv = float(v)
                if not math.isnan(fv):
                    item[key] = fv
            except (TypeError, ValueError):
                pass
        if item.get("plm_MPa") is not None or item.get("EM_MPa") is not None:
            out.append(item)
    return sorted(out, key=lambda x: x["depth_m"])


def resolve_geotech_from_menard(
    embedment_m: float,
    B_m: float,
    gamma_kN_m3: float,
    gamma_buoy_kN_m3: float,
    zw_m: float,
    pressiometer: List[Dict[str, float]],
    q0_manual: Optional[float],
    ple_manual: Optional[float],
    ec_manual: Optional[float],
    ed_manual: Optional[float],
    replacement_thickness_m: float = 0.0,
    replacement_plm_MPa: Optional[float] = None,
    ple_star_zone: str = "mixed",
) -> Dict[str, Any]:
    """合并手动输入与旁压/埋深自动计算。"""
    q0_calc = calc_q0_kpa(embedment_m, gamma_kN_m3, gamma_buoy_kN_m3, zw_m)
    q0 = q0_manual if q0_manual is not None else q0_calc

    zone = ple_star_zone
    if (
        replacement_thickness_m > 0
        and replacement_plm_MPa is not None
        and replacement_plm_MPa > 0
        and not str(ple_star_zone or "").strip()
    ):
        zone = "replacement"

    ple_calc, ple_meta = calc_ple_star_kpa(
        pressiometer, embedment_m, B_m, gamma_kN_m3, gamma_buoy_kN_m3, zw_m,
        replacement_thickness_m, replacement_plm_MPa, zone,
    )
    ec_calc, ed_calc, ed_meta = calc_ec_ed_kpa(pressiometer, embedment_m, B_m)

    ple = ple_manual if ple_manual is not None else ple_calc
    ec = ec_manual if ec_manual is not None else ec_calc
    ed = ed_manual if ed_manual is not None else ed_calc

    return {
        "q0_kPa": round(q0, 2),
        "q0_source": "manual" if q0_manual is not None else "gamma_x_De",
        "q0_calc_kPa": round(q0_calc, 2),
        "ple_star_kPa": ple,
        "ple_source": "manual" if ple_manual is not None else ("menard" if ple_calc else "missing"),
        "ple_meta": ple_meta,
        "Ec_kPa": ec,
        "Ed_kPa": ed,
        "Ec_source": "manual" if ec_manual is not None else ("menard" if ec_calc else "missing"),
        "Ed_source": "manual" if ed_manual is not None else ("menard" if ed_calc else "missing"),
        "menard_layers": ed_meta.get("layers_kPa", {}),
    }
