"""从内力提取 Excel 生成 expert_M 用的 input_NM（纵向 / 横向）。"""
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

from aggregate import ForceCell, StationForces, build_force_tables

INPUT_NM_COLUMNS = [
    "distance_m", "ID", "b_cm", "h_cm", "d1_cm", "d2_cm",
    "is_column", "load_type", "N_kN", "M_kNm",
]

# 排序：里程 → 横向位置（纵向=0）→ ELU 先于 ELS → max 先于 min
_WORD_ORDER = {"ELU": 0, "ELS": 1}
_ENVELOPE_ORDER = {"max": 0, "min": 1}


def _expert_load_type(word: str) -> str:
    """Word 组合 ELU/ELS → Expert 荷载类型 ULS/SLS。"""
    return "ULS" if word == "ELU" else "SLS"


def _station_base(st, is_column: int) -> dict:
    return {
        "distance_m": round(st.global_m, 4),
        "ID": st.french_id(),
        "b_cm": st.b_cm,
        "h_cm": st.h_cm,
        "d1_cm": st.d1_cm,
        "d2_cm": st.d2_cm,
        "is_column": is_column,
        "N_kN": 0.0,
    }


def _append_envelope_rows(
    rows: List[dict],
    base: dict,
    word: str,
    cell: ForceCell,
    trans_idx: int = 0,
) -> None:
    """每个 ELU/ELS 组合最多 2 行：max、min（与 Word 弯矩表一致）。"""
    lt = _expert_load_type(word)
    sort_trans = trans_idx
    if cell.max_val is not None and not pd.isna(cell.max_val):
        rows.append({
            **base,
            "load_type": lt,
            "M_kNm": float(cell.max_val),
            "_sort_word": _WORD_ORDER.get(word, 9),
            "_sort_env": _ENVELOPE_ORDER["max"],
            "_sort_trans": sort_trans,
        })
    if cell.min_val is not None and not pd.isna(cell.min_val):
        rows.append({
            **base,
            "load_type": lt,
            "M_kNm": float(cell.min_val),
            "_sort_word": _WORD_ORDER.get(word, 9),
            "_sort_env": _ENVELOPE_ORDER["min"],
            "_sort_trans": sort_trans,
        })


def _finalize_df(rows: List[dict]) -> pd.DataFrame:
    if not rows:
        raise ValueError("未生成任何 input_NM 行，请检查荷载映射与内力 sheet")
    out = pd.DataFrame(rows)
    out = out.sort_values(
        ["distance_m", "_sort_trans", "_sort_word", "_sort_env"],
        kind="stable",
    ).reset_index(drop=True)
    return out.drop(columns=["_sort_word", "_sort_env", "_sort_trans"])[INPUT_NM_COLUMNS]


def build_input_nm_longitudinal(
    excel_path: Path,
    mapping_path: Path,
    normal_spacing: float = 1.0,
    is_column: int = 0,
) -> pd.DataFrame:
    """
    纵向 input_NM：M ← Mxx 包络（各横向位置取最不利后合并，同 Word 纵向列）。
    每个 ID 4 组荷载：ELU max/min + ELS max/min；N 恒为 0。
    """
    _, moment_rows, _ = build_force_tables(
        excel_path, mapping_path, normal_spacing, mode="moment",
    )
    rows: List[dict] = []
    for sf in moment_rows:
        base = _station_base(sf.station, is_column)
        for word in ("ELU", "ELS"):
            combo = sf.combos.get(word, {})
            cell = combo.get("long")
            if cell:
                _append_envelope_rows(rows, base, word, cell, trans_idx=0)
    return _finalize_df(rows)


def build_input_nm_transverse(
    excel_path: Path,
    mapping_path: Path,
    normal_spacing: float = 1.0,
    is_column: int = 0,
) -> pd.DataFrame:
    """
    横向 input_NM：M ← Myy，按「内力提取1/2/…」分列包络。
    每个 ID = 4 × 横向表个数 组荷载（每横向位置 ELU max/min + ELS max/min）；N 恒为 0。
    """
    _, moment_rows, n_trans = build_force_tables(
        excel_path, mapping_path, normal_spacing, mode="moment",
    )
    rows: List[dict] = []
    for sf in moment_rows:
        base = _station_base(sf.station, is_column)
        for word in ("ELU", "ELS"):
            combo = sf.combos.get(word, {})
            for ti in range(1, n_trans + 1):
                cell = combo.get(f"trans_{ti}")
                if cell:
                    _append_envelope_rows(rows, base, word, cell, trans_idx=ti)
    return _finalize_df(rows)


def export_input_nm(
    excel_path: Path,
    mapping_path: Path,
    output_dir: Path,
    normal_spacing: float = 1.0,
) -> Dict[str, Path]:
    """导出 input_NM_纵向.xlsx 与 input_NM_横向.xlsx。"""
    output_dir.mkdir(parents=True, exist_ok=True)
    long_df = build_input_nm_longitudinal(excel_path, mapping_path, normal_spacing)
    trans_df = build_input_nm_transverse(excel_path, mapping_path, normal_spacing)
    long_path = output_dir / "input_NM_纵向.xlsx"
    trans_path = output_dir / "input_NM_横向.xlsx"
    long_df.to_excel(long_path, index=False)
    trans_df.to_excel(trans_path, index=False)
    return {"longitudinal": long_path, "transverse": trans_path}
