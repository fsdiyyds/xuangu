"""多横向 sheet 归并 + 按站点/组合整理内力。"""
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

import pandas as pd

from load_classify import classify_load, read_load_mapping
from read_forces import (
    align_elements_to_stations,
    find_force_sheets,
    read_force_sheet,
)
from span_layout import Station, build_stations, parse_span_info


@dataclass
class ForceCell:
    max_val: Optional[float] = None
    min_val: Optional[float] = None


@dataclass
class StationForces:
    station: Station
    combos: Dict[str, Dict[str, ForceCell]] = field(default_factory=dict)
    # combos[word_combo][field_key] -> ForceCell
    # field_key: "long" | "trans_1" | "trans_2" ...


def _update_cell(cell: ForceCell, value: float, envelope: str):
    if pd.isna(value):
        return
    v = float(value)
    if envelope == "MAX":
        cell.max_val = v if cell.max_val is None else max(cell.max_val, v)
    elif envelope == "MIN":
        cell.min_val = v if cell.min_val is None else min(cell.min_val, v)


def _long_agg(values: List[float], envelope: str) -> Optional[float]:
    vals = [v for v in values if v is not None and not pd.isna(v)]
    if not vals:
        return None
    return max(vals) if envelope == "MAX" else min(vals)


def build_force_tables(
    excel_path,
    mapping_path,
    normal_spacing: float = 1.0,
    mode: str = "moment",
) -> Tuple[List[Station], List[StationForces], int]:
    """
    mode: moment -> Mxx 纵向, Myy 横向
          shear  -> Vxx 纵向, Vyy 横向
    """
    mapping = read_load_mapping(mapping_path)
    xl = pd.ExcelFile(excel_path)
    if "跨度信息" not in xl.sheet_names:
        raise ValueError("Excel 缺少 sheet「跨度信息」")

    span_df = pd.read_excel(excel_path, sheet_name="跨度信息", dtype=object)
    main_spans, start_len, end_len, span_dims = parse_span_info(span_df)
    stations = build_stations(
        main_spans, start_len, end_len, normal_spacing, span_dims,
    )
    n_st = len(stations)

    force_sheets = find_force_sheets(xl)
    if not force_sheets:
        raise ValueError("未找到名为「内力提取1」「内力提取2」… 的 sheet")

    long_col = "mxx" if mode == "moment" else "vxx"
    trans_col = "myy" if mode == "moment" else "vyy"

    # station_idx -> word_combo -> {long: (max,min), trans_i: (max,min)}
    raw: List[Dict] = [
        {} for _ in range(n_st)
    ]

    for si, sheet in enumerate(force_sheets):
        df = read_force_sheet(excel_path, sheet)
        # 每个横向 sheet 独立：本 sheet 内单元号升序 → 纵桥向站点 0..n-1
        elem_map = align_elements_to_stations(df["element"].tolist(), n_st, sheet)

        for _, row in df.iterrows():
            eid = row["element"]
            if eid not in elem_map:
                continue
            st_idx = elem_map[eid]
            cat, envelope, word = classify_load(row["load"], mapping)
            if not word or not envelope:
                continue

            bucket = raw[st_idx].setdefault(word, {
                "long_max": [], "long_min": [],
                "trans": {},
            })
            trans_key = f"trans_{si + 1}"
            tb = bucket["trans"].setdefault(trans_key, {"max": [], "min": []})

            lv = row[long_col]
            tv = row[trans_col]
            if not pd.isna(lv):
                if envelope == "MAX":
                    bucket["long_max"].append(float(lv))
                else:
                    bucket["long_min"].append(float(lv))
            if not pd.isna(tv):
                if envelope == "MAX":
                    tb["max"].append(float(tv))
                else:
                    tb["min"].append(float(tv))

    n_trans = len(force_sheets)
    result: List[StationForces] = []
    for i, st in enumerate(stations):
        sf = StationForces(station=st)
        for word in ("ELU", "ELS"):
            data = raw[i].get(word)
            if not data:
                sf.combos[word] = {}
                continue
            fields: Dict[str, ForceCell] = {}
            fields["long"] = ForceCell(
                max_val=_long_agg(data["long_max"], "MAX"),
                min_val=_long_agg(data["long_min"], "MIN"),
            )
            for tk in sorted(data["trans"].keys(), key=lambda x: int(x.split("_")[1])):
                td = data["trans"][tk]
                fields[tk] = ForceCell(
                    max_val=_long_agg(td["max"], "MAX"),
                    min_val=_long_agg(td["min"], "MIN"),
                )
            sf.combos[word] = fields
        result.append(sf)

    return stations, result, n_trans


def fmt_val(v: Optional[float], digits: int = 1) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return f"{float(v):.{digits}f}"
