# -*- coding: utf-8 -*-
"""
Midas Civil 内力粘贴 → 浅基础荷载工况自动分类

读取 load_mapping.txt，将「荷载」列映射为 ULS / ALS / SLS（及 QP / RARE），
并转换为 ELU / ELS 承载力、滑动、配筋所需 N、ex、ey、Hd 等。
"""

from __future__ import annotations

import math
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_MAPPING_PATH = SCRIPT_DIR / "load_mapping.txt"

MIDAS_SECTION_KEY = "Midas_Forces"

MIDAS_COLUMNS = [
    "element", "load_case", "location",
    "N_kN", "shear_y_kN", "shear_z_kN", "torsion_kNm",
    "moment_y_kNm", "moment_z_kNm",
]

# 表头别名（Midas 中文导出 / 英文列名）
_HEADER_ALIASES: Dict[str, str] = {
    "element": "element",
    "单元": "element",
    "load_case": "load_case",
    "荷载": "load_case",
    "location": "location",
    "位置": "location",
    "n_kn": "N_kN",
    "轴向": "N_kN",
    "轴向 (kn)": "N_kN",
    "shear_y_kn": "shear_y_kN",
    "剪力-y": "shear_y_kN",
    "剪力-y (kn)": "shear_y_kN",
    "shear_z_kn": "shear_z_kN",
    "剪力-z": "shear_z_kN",
    "剪力-z (kn)": "shear_z_kN",
    "torsion_knm": "torsion_kNm",
    "扭矩": "torsion_kNm",
    "扭矩 (kn·m)": "torsion_kNm",
    "moment_y_knm": "moment_y_kNm",
    "弯矩-y": "moment_y_kNm",
    "弯矩-y (kn·m)": "moment_y_kNm",
    "moment_z_knm": "moment_z_kNm",
    "弯矩-z": "moment_z_kNm",
    "弯矩-z (kn·m)": "moment_z_kNm",
}

_CATEGORY_TO_REGIME = {
    "ULS": "ELU",
    "ALS": "ELUA",
    "SLS": "ELS",
}


def _norm_header(s: str) -> str:
    return re.sub(r"\s+", "", str(s).strip().lower())


def _f(val, default: float = 0.0) -> float:
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return default
    try:
        return float(val)
    except (TypeError, ValueError):
        return default


def _normalize_load_key(name: str) -> str:
    """ELU1(最大) 与 Excel 乱码 ELU1(??) 均归一为 ELU1。"""
    s = str(name).strip()
    if "(" in s:
        s = s.split("(", 1)[0].strip()
    return s


def _normalize_category(cat: str) -> str:
    """映射表可用 ELU/ELS/ELA 或 ULS/SLS/ALS。"""
    c = re.sub(r"\s+", "", str(cat).strip().upper())
    if c in ("ELU", "ULS"):
        return "ULS"
    if c in ("ELUA", "ELA", "ALS"):
        return "ALS"
    if c in ("ELS", "SLS"):
        return "SLS"
    if "QP" in c:
        return "SLS"
    if "RARE" in c or "RARES" in c:
        return "SLS"
    return c


def _infer_sls_subtype(name: str) -> Optional[str]:
    upper = name.upper()
    nkey = _normalize_load_key(name).upper()
    if "RARE" in upper or "RARES" in nkey or "稀有" in name:
        return "RARE"
    if "QP" in upper or "准永久" in name or "长期" in name or "QUASI" in upper:
        return "QP"
    if (
        "FRÉQUENT" in upper or "FREQUENT" in upper or "FREQUENTES" in upper
        or "频繁" in name
    ):
        return "FREQ"
    return "FREQ"


def _normalize_subtype(
    cat: str, subtype: Optional[str], load_name: str,
) -> Optional[str]:
    parts = [str(subtype or ""), str(cat)]
    blob = " ".join(parts).upper()
    blob_ns = blob.replace(" ", "")
    if "QP" in blob_ns:
        return "QP"
    if "RARE" in blob or "RARES" in blob:
        return "RARE"
    if (
        "FRÉQUENT" in blob or "FREQUENT" in blob
        or "FREQUENTES" in blob or "频繁" in blob
    ):
        return "FREQ"
    return _infer_sls_subtype(load_name)


def _finalize_classification(
    cat: str, subtype: Optional[str], load_name: str,
) -> Tuple[str, Optional[str]]:
    norm_cat = _normalize_category(cat)
    if norm_cat == "SLS":
        st = _normalize_subtype(cat, subtype, load_name)
        return norm_cat, st
    return norm_cat, None


def load_mapping_file(path: Optional[Path] = None) -> List[Tuple[str, str, Optional[str]]]:
    """
    解析映射表，每行：荷载名称, ULS|ALS|SLS [, QP|RARE]
    """
    p = path or DEFAULT_MAPPING_PATH
    rules: List[Tuple[str, str, Optional[str]]] = []
    if not p.is_file():
        return rules
    for line in p.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" in line and "," not in line:
            name, cat = line.split("=", 1)
            parts = [name.strip(), cat.strip()]
        else:
            parts = [x.strip() for x in line.split(",")]
            name = parts[0]
            cat = parts[1] if len(parts) > 1 else ""
        subtype = None
        if len(parts) > 2 and parts[2]:
            subtype = parts[2].upper()
        cat = cat.strip().upper()
        if name and cat:
            rules.append((name.strip(), cat, subtype))
    return rules


def classify_load_case(
    load_name: str,
    rules: List[Tuple[str, str, Optional[str]]],
) -> Tuple[Optional[str], Optional[str]]:
    """返回 (ULS|ALS|SLS, QP|RARE|None)。"""
    name = str(load_name).strip()
    if not name:
        return None, None
    nkey = _normalize_load_key(name)

    for key, cat, subtype in rules:
        if name == key or nkey == _normalize_load_key(key):
            return _finalize_classification(cat, subtype, name)

    sorted_rules = sorted(rules, key=lambda x: len(x[0]), reverse=True)
    for key, cat, subtype in sorted_rules:
        knorm = _normalize_load_key(key)
        if (
            name.startswith(key)
            or nkey.startswith(knorm)
            or (len(knorm) >= 3 and nkey == knorm)
        ):
            return _finalize_classification(cat, subtype, name)

    # 启发式
    upper = name.upper()
    nupper = nkey.upper()
    if "ELUA" in upper or "ACCIDENT" in upper or "地震" in name:
        return "ALS", None
    if nupper.startswith("ELU") or "ULS" in upper:
        return "ULS", None
    if "ELS" in upper or "SLS" in upper:
        return "SLS", _infer_sls_subtype(name)
    return None, None


def normalize_midas_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """将粘贴表头统一为 MIDAS_COLUMNS。"""
    if df.empty:
        return pd.DataFrame(columns=MIDAS_COLUMNS)

    col_map: Dict[str, str] = {}
    for c in df.columns:
        key = _norm_header(c)
        for alias, std in _HEADER_ALIASES.items():
            if _norm_header(alias) == key or key in _norm_header(alias):
                col_map[c] = std
                break
        if c not in col_map and key in _HEADER_ALIASES:
            col_map[c] = _HEADER_ALIASES[key]

    out = pd.DataFrame()
    for std in MIDAS_COLUMNS:
        src = next((orig for orig, s in col_map.items() if s == std), None)
        if src is not None:
            out[std] = df[src]
        else:
            out[std] = None
    return out


def _norm_elem_id(val: Any) -> str:
    s = str(val).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s


def _parse_element_filter(raw: Any) -> Optional[set]:
    if raw is None or (isinstance(raw, float) and math.isnan(raw)):
        return None
    s = str(raw).strip()
    if not s:
        return None
    ids = set()
    for part in re.split(r"[,;，；\s]+", s):
        part = _norm_elem_id(part)
        if part:
            ids.add(part)
    return ids or None


def _row_to_case(
    row: Dict[str, Any],
    B_m: float,
    L_m: float,
) -> Dict[str, Any]:
    """单行 Midas 内力 → 基础验算荷载。"""
    n_raw = _f(row.get("N_kN"))
    n_kn = abs(n_raw)
    my = _f(row.get("moment_y_kNm"))
    mz = _f(row.get("moment_z_kNm"))
    vy = _f(row.get("shear_y_kN"))
    vz = _f(row.get("shear_z_kN"))

    ex_m = abs(my) / n_kn if n_kn > 1.0 else 0.0
    ey_m = abs(mz) / n_kn if n_kn > 1.0 else 0.0
    hd = math.hypot(vy, vz)

    load_case = str(row.get("load_case", "")).strip()
    element = str(row.get("element", "")).strip()
    location = str(row.get("location", "")).strip()

    return {
        "load_case": load_case,
        "element": element,
        "location": location,
        "N_kN": round(n_kn, 2),
        "ex_m": round(ex_m, 4),
        "ey_m": round(ey_m, 4),
        "Hd_kN": round(hd, 2),
        "moment_y_kNm": round(my, 2),
        "moment_z_kNm": round(mz, 2),
        "Mmax_ELU_kNm_m": round(abs(my) / L_m, 2) if L_m > 0 else round(abs(my), 2),
        "Mmax_ELS_kNm_m": round(abs(my) / L_m, 2) if L_m > 0 else round(abs(my), 2),
        "Vd_ELU_kN": round(hd, 2),
    }


def _extract_elu_raw_rows(
    df: pd.DataFrame,
    rules: List[Tuple[str, str, Optional[str]]],
    element_filter: Optional[set],
) -> List[Dict[str, Any]]:
    """保留 ELU 原始 Midas 行（供 Fig 8.2.3.1 截面力计算）。"""
    return _extract_midas_raw_rows(df, rules, element_filter, ("ULS", "ALS"))


def _extract_els_frequent_raw_rows(
    df: pd.DataFrame,
    rules: List[Tuple[str, str, Optional[str]]],
    element_filter: Optional[set],
) -> List[Dict[str, Any]]:
    """保留 ELS fréquentes 原始 Midas 行（A-A/B-B 弯矩 ELS 包络）。"""
    rows: List[Dict[str, Any]] = []
    for _, r in df.iterrows():
        row = {c: r.get(c) for c in MIDAS_COLUMNS}
        if all(pd.isna(v) or str(v).strip() == "" for v in row.values()):
            continue
        lc = str(row.get("load_case", "")).strip()
        if not lc or lc in {"荷载", "列名（勿改）"}:
            continue
        cat, subtype = classify_load_case(lc, rules)
        if cat != "SLS" or subtype in ("QP", "RARE"):
            continue
        elem = _norm_elem_id(row.get("element", ""))
        if element_filter and elem not in element_filter:
            continue
        rows.append({
            "element": elem,
            "load_case": lc,
            "N_kN": _f(row.get("N_kN")),
            "shear_y_kN": _f(row.get("shear_y_kN")),
            "shear_z_kN": _f(row.get("shear_z_kN")),
            "moment_y_kNm": _f(row.get("moment_y_kNm")),
            "moment_z_kNm": _f(row.get("moment_z_kNm")),
        })
    return rows


def _extract_midas_raw_rows(
    df: pd.DataFrame,
    rules: List[Tuple[str, str, Optional[str]]],
    element_filter: Optional[set],
    categories: Tuple[str, ...],
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    for _, r in df.iterrows():
        row = {c: r.get(c) for c in MIDAS_COLUMNS}
        if all(pd.isna(v) or str(v).strip() == "" for v in row.values()):
            continue
        lc = str(row.get("load_case", "")).strip()
        if not lc or lc in {"荷载", "列名（勿改）"}:
            continue
        cat, _ = classify_load_case(lc, rules)
        if cat not in categories:
            continue
        elem = _norm_elem_id(row.get("element", ""))
        if element_filter and elem not in element_filter:
            continue
        rows.append({
            "element": elem,
            "load_case": lc,
            "N_kN": _f(row.get("N_kN")),
            "shear_y_kN": _f(row.get("shear_y_kN")),
            "shear_z_kN": _f(row.get("shear_z_kN")),
            "moment_y_kNm": _f(row.get("moment_y_kNm")),
            "moment_z_kNm": _f(row.get("moment_z_kNm")),
        })
    return rows


def build_load_tables_from_midas(
    midas_df: pd.DataFrame,
    element_filter: Optional[set],
    B_m: float,
    L_m: float,
    mapping_path: Optional[Path] = None,
) -> Dict[str, List[Dict[str, Any]]]:
    """
    从 Midas 表生成各验算区块荷载列表。
    """
    rules = load_mapping_file(mapping_path)
    df = normalize_midas_dataframe(midas_df)
    elu: List[Dict[str, Any]] = []
    els_frequent: List[Dict[str, Any]] = []
    els_qp: List[Dict[str, Any]] = []
    els_rare: List[Dict[str, Any]] = []
    unmapped: List[str] = []

    for _, r in df.iterrows():
        row = {c: r.get(c) for c in MIDAS_COLUMNS}
        if all(
            pd.isna(v) or str(v).strip() == ""
            for v in row.values()
        ):
            continue
        elem = _norm_elem_id(row.get("element", ""))
        if element_filter and elem not in element_filter:
            continue

        load_name = str(row.get("load_case", "")).strip()
        if load_name in {"荷载", "单元", "位置", "列名（勿改）", "填写说明 →", "单位"}:
            continue
        cat, subtype = classify_load_case(load_name, rules)
        if cat is None:
            if load_name not in unmapped:
                unmapped.append(load_name)
            continue

        case = _row_to_case(row, B_m, L_m)
        case["limit_state"] = cat
        case["sls_subtype"] = subtype or ""

        if cat in ("ULS", "ALS"):
            elu.append(case)
        elif cat == "SLS":
            if subtype == "QP":
                els_qp.append(case)
            elif subtype == "RARE":
                els_rare.append(case)
            else:
                # ELS fréquentes / 常规 ELS → 沉降 Evaluation du tassement
                els_frequent.append(case)

    hd_only = [{"Hd_kN": c["Hd_kN"], "load_case": c.get("load_case", "")} for c in elu]

    return {
        "ELU_mobilisation": elu,
        "ELS_mobilisation_FREQ": els_frequent,
        "ELS_mobilisation_QP": els_qp,
        "ELS_mobilisation_RARE": els_rare,
        "ELS_mobilisation": els_frequent,
        "ELU_glissement": hd_only,
        "_midas_unmapped": unmapped,
        "_midas_source": True,
        "_midas_elu_raw": _extract_elu_raw_rows(df, rules, None),
        "_midas_els_freq_raw": _extract_els_frequent_raw_rows(df, rules, None),
    }


def resolve_pier_loads(
    pier_sections: Dict[str, pd.DataFrame],
    pier_params: Dict[str, Any],
    B_m: float,
    L_m: float,
    mapping_path: Optional[Path] = None,
) -> Dict[str, Any]:
    """
    优先 Midas 粘贴区；否则读取手工 ELU/ELS 表格（向后兼容）。
    """
    midas_df = pier_sections.get(MIDAS_SECTION_KEY, pd.DataFrame())
    norm = normalize_midas_dataframe(midas_df) if not midas_df.empty else pd.DataFrame()
    has_midas = False
    if not norm.empty:
        for _, r in norm.iterrows():
            lc = str(r.get("load_case", "")).strip()
            if lc and lc not in {"荷载", "列名（勿改）"}:
                has_midas = True
                break

    if has_midas:
        elem_raw = pier_params.get("midas_element_ids")
        filt = _parse_element_filter(elem_raw)
        tables = build_load_tables_from_midas(
            midas_df, filt, B_m, L_m, mapping_path,
        )
        meta = {
            "load_source": "midas",
            "midas_unmapped": tables.get("_midas_unmapped", []),
            "elu_count": len(tables["ELU_mobilisation"]),
            "els_freq_count": len(tables.get("ELS_mobilisation_FREQ", [])),
            "els_qp_count": len(tables["ELS_mobilisation_QP"]),
            "els_rare_count": len(tables["ELS_mobilisation_RARE"]),
        }
        tables["load_meta"] = meta
        return tables

    def _read(key: str) -> List[Dict[str, Any]]:
        df = pier_sections.get(key, pd.DataFrame())
        if df.empty:
            return []
        rows: List[Dict[str, Any]] = []
        for _, row in df.iterrows():
            d = {str(c).strip(): row[c] for c in df.columns}
            rows.append(d)
        return rows

    elu = _read("ELU_mobilisation")
    els = _read("ELS_mobilisation")
    return {
        "ELU_mobilisation": elu,
        "ELS_mobilisation_FREQ": els,
        "ELS_mobilisation_QP": [],
        "ELS_mobilisation_RARE": els,
        "ELS_mobilisation": els,
        "ELU_glissement": _read("ELU_glissement"),
        "load_meta": {"load_source": "manual"},
        "_midas_elu_raw": [],
        "_midas_els_freq_raw": [],
    }
