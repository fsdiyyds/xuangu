# -*- coding: utf-8 -*-
"""
Fascicule 62 Titre V — Annexe F.1 : coefficient minorateur iδβ

Charge centrée inclinée sur sol horizontal (§2) :
  · Sols cohérents (argiles, limons…) : iδβ = Φ1(δ)
  · Sols frottants (sables, graves…)  : iδβ = Φ2(δ, De, B)

δ : inclinaison absolue de la charge par rapport à la verticale (rad).

Pour charge verticale avec excentricité ex (sans Hd) : on prend l'inclinaison
équivalente de contrainte δ = arctan(2·ex/B) (NF P94-261 / Meyerhof, petits angles).

Sols mixtes : interpolation Φ1–Φ2 selon NF P94-261 (α=0,6).
"""

from __future__ import annotations

import math
from typing import Optional, Tuple

SOIL_COHERENT = "coherent"      # argiles, limons, craies…
SOIL_FRICTION = "frictional"    # sables, graves
SOIL_MIXED = "mixed"


def classify_soil(phi_deg: float, C_kPa: float, soil_category: Optional[str] = None) -> str:
    """判定土质类别（与 Annexe E.1 / NF P94-261 一致）。"""
    if soil_category:
        s = str(soil_category).strip().lower()
        mapping = {
            "argile": SOIL_COHERENT, "coherent": SOIL_COHERENT, "cohérent": SOIL_COHERENT,
            "sable": SOIL_FRICTION, "sableux": SOIL_FRICTION, "frictional": SOIL_FRICTION,
            "frottant": SOIL_FRICTION, "grave": SOIL_FRICTION,
            "mixte": SOIL_MIXED, "mixed": SOIL_MIXED,
        }
        for key, val in mapping.items():
            if key in s:
                return val
    if phi_deg > 5 and C_kPa > 5:
        return SOIL_MIXED
    if phi_deg > 5 and C_kPa <= 5:
        return SOIL_FRICTION
    return SOIL_COHERENT


def phi1(delta_rad: float) -> float:
    """Φ1(δ) = (1 − 2δ/π)² — sols cohérents."""
    t = 2.0 * delta_rad / math.pi
    return max(0.0, (1.0 - t) ** 2)


def phi2(delta_rad: float, De_m: float, B_m: float) -> float:
    """
    Φ2(δ) — sols frottants (Fasc. 62 Annexe F.1 §2.2).
    δ en radians ; De = encastrement équivalent ; B = largeur.
    """
    if B_m <= 0:
        return phi1(delta_rad)
    t = 2.0 * delta_rad / math.pi
    base = (1.0 - t) ** 2
    exp_term = math.exp(-De_m / B_m)
    if delta_rad < math.pi / 4.0:
        corr = t * (2.0 - 3.0 * t) * exp_term
    else:
        corr = base * exp_term
    return max(0.0, base - corr)


def delta_rad_from_load(
    ex_m: float,
    ey_m: float,
    B_m: float,
    L_m: float,
    Hd_kN: Optional[float] = None,
    N_kN: Optional[float] = None,
) -> float:
    """
    δ absolu (rad) :
      - si Hd, N > 0 : arctan(Hd/N)
      - sinon excentricité : arctan(2·ex/B) (ex 沿 B)
    """
    if Hd_kN is not None and N_kN is not None and N_kN > 0 and abs(Hd_kN) > 1e-6:
        return math.atan(abs(Hd_kN) / abs(N_kN))
    ex_m = max(ex_m, 0.0)
    if B_m > 0 and ex_m > 0:
        return math.atan(2.0 * ex_m / B_m)
    if L_m > 0 and ey_m > 0 and B_m > 0:
        return math.atan(2.0 * ey_m / L_m)
    return 0.0


def calc_i_delta_beta(
    ex_m: float,
    B_m: float,
    De_m: float,
    soil_class: str,
    phi_deg: float = 25.0,
    C_kPa: float = 0.0,
    ey_m: float = 0.0,
    L_m: float = 0.0,
    Hd_kN: Optional[float] = None,
    N_kN: Optional[float] = None,
) -> Tuple[float, float, float, str]:
    """
    返回 (iδβ, δ_deg, phi_used, formula_note).
    """
    delta_r = delta_rad_from_load(ex_m, ey_m, B_m, L_m, Hd_kN, N_kN)
    delta_deg = math.degrees(delta_r)
    p1 = phi1(delta_r)
    p2 = phi2(delta_r, De_m, B_m)

    if soil_class == SOIL_COHERENT:
        i_beta = p1
        note = "argile/coherent: iδβ=Φ1(δ)=(1−2δ/π)²"
        phi_used = p1
    elif soil_class == SOIL_FRICTION:
        i_beta = p2
        note = "sable/frottant: iδβ=Φ2(δ,De,B)"
        phi_used = p2
    else:
        # NF P94-261 : Φ2 + (Φ1−Φ2)(1−exp(−α·c/(γ·B·tanφ)))，α=0,6
        phi_rad = math.radians(phi_deg)
        tan_phi = math.tan(phi_rad) if phi_deg > 0 else 1.0
        gamma = 20.0
        alpha = 0.6
        if C_kPa > 0 and B_m > 0 and tan_phi > 0:
            blend = 1.0 - math.exp(-alpha * C_kPa / (gamma * B_m * tan_phi))
        else:
            blend = 0.0
        i_beta = p2 + (p1 - p2) * blend
        note = "mixte: Φ2+(Φ1−Φ2)·(1−exp(−0,6c/(γBtanφ)))"
        phi_used = i_beta

    i_beta = max(0.0, min(1.0, i_beta))
    return round(i_beta, 4), round(delta_deg, 3), round(phi_used, 4), note
