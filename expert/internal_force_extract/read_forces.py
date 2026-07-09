"""读取「内力提取」sheet 并标准化列名。"""
import re
from typing import Dict, List, Optional, Tuple

import pandas as pd

FORCE_SHEET_PATTERN = re.compile(r"^内力提取\s*\d+$", re.I)

COL_ALIASES = {
    "element": ("单元", "element", "elem", "板单元"),
    "load": ("荷载", "load", "case", "工况"),
    "mxx": ("mxx",),
    "myy": ("myy",),
    "fxx": ("fxx",),
    "vxx": ("vxx",),
    "vyy": ("vyy",),
    "fyy": ("fyy",),
}


def _norm_header(h) -> str:
    s = str(h).strip().lower()
    s = s.replace("（", "(").replace("）", ")")
    s = re.sub(r"\s+", "", s)
    return s


def _match_col(columns, aliases) -> Optional[str]:
    for c in columns:
        nh = _norm_header(c)
        for a in aliases:
            if nh == _norm_header(a) or nh.startswith(_norm_header(a)):
                return c
        if "mxx" in aliases and "mxx" in nh:
            return c
        if "myy" in aliases and "myy" in nh:
            return c
        if "fxx" in aliases and nh.startswith("fxx"):
            return c
        if "vxx" in aliases and nh.startswith("vxx"):
            return c
        if "fyy" in aliases and nh.startswith("fyy"):
            return c
        if "vyy" in aliases and nh.startswith("vyy"):
            return c
    return None


def normalize_force_df(df: pd.DataFrame) -> pd.DataFrame:
    cols = list(df.columns)
    mapping = {}
    for std, aliases in COL_ALIASES.items():
        c = _match_col(cols, aliases)
        if c:
            mapping[c] = std
    if "element" not in mapping.values() or "load" not in mapping.values():
        raise ValueError(f"内力 sheet 缺少「单元」或「荷载」列，当前列: {cols}")

    out = df.rename(columns={k: v for k, v in mapping.items()}).copy()
    for c in ("mxx", "myy", "fxx", "vxx", "vyy"):
        if c not in out.columns:
            out[c] = float("nan")
        out[c] = pd.to_numeric(out[c], errors="coerce")

    out["element"] = out["element"].apply(_norm_element_id)
    out = out[out["element"].astype(str).str.len() > 0].copy()
    return out[["element", "load", "mxx", "myy", "fxx", "vxx", "vyy"]]


def _norm_element_id(val) -> str:
    if pd.isna(val):
        return ""
    s = str(val).strip()
    try:
        f = float(s)
        if f == int(f):
            return str(int(f))
    except (ValueError, TypeError):
        pass
    return s


def find_force_sheets(xl: pd.ExcelFile) -> List[str]:
    return sorted(
        [s for s in xl.sheet_names if FORCE_SHEET_PATTERN.match(str(s).strip())],
        key=lambda x: int(re.search(r"\d+", x).group()),
    )


def read_force_sheet(path, sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet_name, dtype=object)
    return normalize_force_df(df)


def _sort_element_ids(element_ids: List[str]) -> List[str]:
    """同一 sheet 内去重后按单元号从小到大排序（纵桥向里程顺序）。"""
    uniq = {e for e in element_ids if e}
    return sorted(uniq, key=lambda x: (0, int(x)) if str(x).isdigit() else (1, str(x)))


def align_elements_to_stations(
    element_ids: List[str], n_stations: int, sheet_name: str,
) -> Dict[str, int]:
    """
    单个「内力提取」sheet 内：单元号升序 ↔ 纵桥向站点索引。

    各横向 sheet 的单元编号彼此独立（可完全不同），但每个 sheet 内
    单元号从小到大即沿桥纵向排列，第 i 小对应第 i 个站点（含首末支座段）。
    """
    ordered = _sort_element_ids(element_ids)
    if len(ordered) != n_stations:
        preview = ", ".join(ordered[:3] + (["…"] if len(ordered) > 6 else []) + ordered[-3:])
        raise ValueError(
            f"{sheet_name}: 单元数 {len(ordered)} 与跨度信息站点数 {n_stations} 不一致。"
            f"（升序单元示例: {preview}）请检查跨度、间距或该 sheet 数据。"
        )
    return {e: i for i, e in enumerate(ordered)}
