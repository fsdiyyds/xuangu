"""从「跨度信息」sheet 生成纵桥向里程与 Travée 编号。"""
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

import pandas as pd

DEFAULT_DIMS = {"b_cm": 100.0, "h_cm": 126.2, "d1_cm": 5.0, "d2_cm": 5.0}


@dataclass
class Station:
    travee: int
    local_m: float
    position_label: str
    global_m: float
    is_start_bearing: bool = False
    is_end_bearing: bool = False
    b_cm: float = DEFAULT_DIMS["b_cm"]
    h_cm: float = DEFAULT_DIMS["h_cm"]
    d1_cm: float = DEFAULT_DIMS["d1_cm"]
    d2_cm: float = DEFAULT_DIMS["d2_cm"]

    def french_id(self) -> str:
        """ID 法语：Travée + 距离（支座段单独标注）。"""
        if self.is_start_bearing:
            return f"Travée {self.travee} + appui 0,5 m (début)"
        if self.is_end_bearing:
            return f"Travée {self.travee} + appui 0,5 m (fin)"
        if abs(self.local_m - round(self.local_m)) < 1e-3:
            dist = f"{int(round(self.local_m))} m"
        else:
            dist = f"{self.local_m:g} m".replace(".", ",")
        return f"Travée {self.travee} + {dist}"


def _find_col(df: pd.DataFrame, *keywords) -> Optional[str]:
    for c in df.columns:
        s = str(c).strip()
        for kw in keywords:
            if kw in s:
                return c
    return None


def _float_or_default(val, default: float) -> float:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return default
    try:
        return float(val)
    except (TypeError, ValueError):
        return default


def parse_span_info(
    df: pd.DataFrame,
) -> Tuple[List[Tuple[str, float]], float, float, Dict[str, Dict[str, float]]]:
    """
    解析跨度信息表。
    返回 (主跨列表, 起点支座长度, 终点支座长度, 各跨尺寸 {跨号: {b_cm,...}})。
    「起点」尺寸并入 1 号跨，「终点」尺寸并入最后一跨。
    """
    name_col = _find_col(df, "桥梁跨度", "跨度")
    len_col = _find_col(df, "桥跨长度", "长度")
    if not name_col or not len_col:
        raise ValueError("「跨度信息」需含列：桥梁跨度、桥跨长度")

    dim_cols = {
        k: _find_col(df, k, k.replace("_", ""))
        for k in ("b_cm", "h_cm", "d1_cm", "d2_cm")
    }

    main_spans: List[Tuple[str, float]] = []
    start_len = 0.5
    end_len = 0.5
    span_dims: Dict[str, Dict[str, float]] = {}
    start_dims = dict(DEFAULT_DIMS)
    end_dims = dict(DEFAULT_DIMS)

    for _, row in df.iterrows():
        label = str(row[name_col]).strip()
        if not label or label.lower() == "nan":
            continue
        try:
            length = float(row[len_col])
        except (TypeError, ValueError):
            continue

        dims = {
            k: _float_or_default(row.get(dim_cols[k]) if dim_cols[k] else None, DEFAULT_DIMS[k])
            for k in DEFAULT_DIMS
        }

        if label in ("起点", "始点"):
            start_len = length
            start_dims = dims
        elif label in ("终点", "结束"):
            end_len = length
            end_dims = dims
        elif label.isdigit():
            main_spans.append((label, length))
            span_dims[label] = dims

    if not main_spans:
        raise ValueError("「跨度信息」中未找到数字编号的桥跨")

    first = main_spans[0][0]
    last = main_spans[-1][0]
    for k, v in start_dims.items():
        if span_dims.get(first, {}).get(k) == DEFAULT_DIMS[k]:
            span_dims.setdefault(first, dict(DEFAULT_DIMS))[k] = v
    for k, v in end_dims.items():
        span_dims.setdefault(last, dict(DEFAULT_DIMS))
        span_dims[last][k] = v

    return main_spans, start_len, end_len, span_dims


def build_stations(
    main_spans: List[Tuple[str, float]],
    start_bearing_len: float = 0.5,
    end_bearing_len: float = 0.5,
    normal_spacing: float = 1.0,
    span_dims: Optional[Dict[str, Dict[str, float]]] = None,
) -> List[Station]:
    """
    纵桥向站点（与板单元号升序一一对应）：
    - 全桥第一个单元：1 号跨起点 0.5 m 支座段（含在 1 号跨内）
    - 全桥最后一个单元：末跨终点 0.5 m 支座段（含在末跨内）
    - 其余按 normal_spacing（默认 1 m）划分
    """
    span_dims = span_dims or {}
    lengths = [L for _, L in main_spans]
    labels = [lab for lab, _ in main_spans]
    n_tr = len(main_spans)
    stations: List[Station] = []

    def _dims_for(t_idx: int) -> Dict[str, float]:
        key = labels[t_idx - 1]
        return {**DEFAULT_DIMS, **span_dims.get(key, {})}

    def _label_m(m: float) -> str:
        if abs(m - round(m)) < 1e-3:
            return f"{int(round(m))}m"
        return f"{m:g}m"

    for t_idx, length in enumerate(lengths, start=1):
        dims = _dims_for(t_idx)
        is_first = t_idx == 1
        is_last = t_idx == n_tr
        seg: List[Tuple[float, bool, bool]] = []

        if is_first:
            seg.append((0.0, True, False))
            pos = 0.0
            while pos + normal_spacing <= length + 1e-6:
                pos = round(pos + normal_spacing, 6)
                seg.append((pos, False, False))
        elif is_last:
            pos = 0.0
            while pos + normal_spacing <= length + 1e-6:
                pos = round(pos + normal_spacing, 6)
                seg.append((pos, False, False))
            seg.append((round(length + end_bearing_len, 6), False, True))
        else:
            pos = 0.0
            while pos + normal_spacing <= length + 1e-6:
                pos = round(pos + normal_spacing, 6)
                seg.append((pos, False, False))

        if is_first:
            prefix_sum = 0.0
        else:
            prefix_sum = sum(lengths[: t_idx - 1])

        for local_m, is_sb, is_eb in seg:
            if is_first:
                global_m = local_m
            else:
                global_m = prefix_sum + local_m
            stations.append(
                Station(
                    travee=t_idx,
                    local_m=local_m,
                    global_m=global_m,
                    position_label=_label_m(global_m),
                    is_start_bearing=is_sb,
                    is_end_bearing=is_eb,
                    **dims,
                )
            )

    return stations
