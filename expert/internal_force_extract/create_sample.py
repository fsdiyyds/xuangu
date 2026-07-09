"""生成与图 4 结构一致的示例 Excel，便于测试。"""
from pathlib import Path

import pandas as pd

from span_layout import build_stations, parse_span_info

_DIM = {"b_cm": 100, "h_cm": 126.2, "d1_cm": 5, "d2_cm": 5}


def write_sample(path: Path):
    span = pd.DataFrame([
        {"桥梁跨度": "起点", "桥跨长度": 0.5, **_DIM},
        {"桥梁跨度": "1", "桥跨长度": 12, **_DIM},
        {"桥梁跨度": "2", "桥跨长度": 15, **_DIM},
        {"桥梁跨度": "3", "桥跨长度": 15, **_DIM},
        {"桥梁跨度": "4", "桥跨长度": 12, **_DIM},
        {"桥梁跨度": "终点", "桥跨长度": 0.5, **_DIM},
    ])
    main_spans, start_len, end_len, span_dims = parse_span_info(span)
    stations = build_stations(
        main_spans, start_len, end_len, normal_spacing=1.0, span_dims=span_dims,
    )
    n = len(stations)
    # 不同横向位置 sheet 使用不同单元号段，但各自升序对应同一纵桥向站点序列
    sheet_start_ids = [6500, 8200]

    loads_max = [
        "ELU包络(最大)", "ELUA包络(最大)", "ELS RARE 包络(最大)",
    ]
    loads_min = [
        "ELU包络(最小)", "ELUA包络(最小)", "ELS RARE 包络(最小)",
    ]

    def make_sheet(start_id: int, n: int, mxx_scale: float, myy_scale: float):
        rows = []
        for i in range(n):
            eid = start_id + i
            for lt in loads_max + loads_min:
                sign = 1 if "最大" in lt else -1
                rows.append({
                    "单元": eid,
                    "荷载": lt,
                    "节点": "中央",
                    "Fxx (kN/m)": round(sign * (10 + i * 2) * mxx_scale, 3),
                    "Fyy (kN/m)": 0,
                    "Mxx (kN·m/m)": round(sign * (50 + i * 20) * mxx_scale, 3),
                    "Myy (kN·m/m)": round(sign * (100 + i * 5) * myy_scale, 3),
                    "Vxx (kN/m)": round(sign * (80 + i * 3) * mxx_scale, 3),
                    "Vyy (kN/m)": round(sign * (90 + i * 4) * myy_scale, 3),
                })
        return pd.DataFrame(rows)

    f1 = make_sheet(sheet_start_ids[0], n, 1.0, 1.0)
    f2 = make_sheet(sheet_start_ids[1], n, 0.9, 1.1)

    with pd.ExcelWriter(path, engine="openpyxl") as w:
        span.to_excel(w, sheet_name="跨度信息", index=False)
        f1.to_excel(w, sheet_name="内力提取1", index=False)
        f2.to_excel(w, sheet_name="内力提取2", index=False)
