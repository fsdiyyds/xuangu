"""
从 Excel Sheet1 读取「区域 / 单元号 / 端」，按 i端、j端 生成 Midas 内力提取用的单元号序列。

Midas 结果表「单元」列通常可直接粘贴：空格分隔的单元号，如
  5949 5950 5951 5952

用法:
  python midas_element_export.py
  python midas_element_export.py "D:\数据\单元表.xlsx"
"""

import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

_SCRIPT_DIR = Path(__file__).resolve().parent

# ===================== 运行配置 =====================
# 留空则运行时输入；默认 demo_midas 内的单元区域表
DEFAULT_EXCEL_PATH = str(_SCRIPT_DIR / "单元区域表.xlsx")
DEFAULT_SHEET_NAME = "Sheet1"
OUTPUT_I_FILE = "midas_units_i端.txt"
OUTPUT_J_FILE = "midas_units_j端.txt"
OUTPUT_DETAIL_FILE = "midas_units_明细.xlsx"
# ==================================================

COL_REGION = "区域"
COL_ELEMENT = "单元号"
COL_END = "端"

HEADER_ALIASES = {
    COL_REGION: ("区域", "region", "部位"),
    COL_ELEMENT: ("单元号", "单元", "单元编号", "element", "elem"),
    COL_END: ("端", "端部", "end", "i/j"),
}


def normalize_text(s) -> str:
    if pd.isna(s):
        return ""
    return re.sub(r"\s+", "", str(s).strip())


def normalize_element_id(value) -> Optional[int]:
    if pd.isna(value):
        return None
    raw = str(value).strip()
    if not raw:
        return None
    try:
        num = float(raw)
        if num == int(num):
            return int(num)
    except (ValueError, TypeError):
        pass
    return None


def normalize_end(value) -> Optional[str]:
    """返回 'i' 或 'j'"""
    if pd.isna(value):
        return None
    s = str(value).strip().lower()
    compact = normalize_text(value).lower()
    if compact in ("i", "i端", "iend") or s.startswith("i"):
        return "i"
    if compact in ("j", "j端", "jend") or s.startswith("j"):
        return "j"
    return None


def _match_column(columns: List[str], aliases: Tuple[str, ...]) -> Optional[str]:
    norm_map = {normalize_text(c).lower(): c for c in columns}
    for alias in aliases:
        key = normalize_text(alias).lower()
        if key in norm_map:
            return norm_map[key]
    for col in columns:
        ncol = normalize_text(col).lower()
        for alias in aliases:
            if normalize_text(alias).lower() in ncol:
                return col
    return None


def read_element_table(excel_path: Path, sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=object)
    df = df.dropna(how="all")
    region_col = _match_column(list(df.columns), HEADER_ALIASES[COL_REGION])
    elem_col = _match_column(list(df.columns), HEADER_ALIASES[COL_ELEMENT])
    end_col = _match_column(list(df.columns), HEADER_ALIASES[COL_END])
    missing = []
    if not region_col:
        missing.append(COL_REGION)
    if not elem_col:
        missing.append(COL_ELEMENT)
    if not end_col:
        missing.append(COL_END)
    if missing:
        raise ValueError(f"未找到列: {missing}，当前表头: {list(df.columns)}")

    out = pd.DataFrame({
        COL_REGION: df[region_col],
        COL_ELEMENT: df[elem_col],
        COL_END: df[end_col],
    })
    out = out.dropna(how="all")
    out[COL_ELEMENT] = out[COL_ELEMENT].apply(normalize_element_id)
    out[COL_END] = out[COL_END].apply(normalize_end)
    out = out.dropna(subset=[COL_ELEMENT, COL_END])
    out[COL_REGION] = out[COL_REGION].astype(str).str.strip()
    return out


def format_midas_sequence(element_ids: List[int]) -> str:
    """Midas 单元列：空格分隔，去重后按单元号排序"""
    unique = sorted(set(element_ids))
    return " ".join(str(n) for n in unique)


def split_by_end(df: pd.DataFrame) -> Tuple[List[int], List[int]]:
    i_ids = df.loc[df[COL_END] == "i", COL_ELEMENT].tolist()
    j_ids = df.loc[df[COL_END] == "j", COL_ELEMENT].tolist()
    return i_ids, j_ids


def build_region_detail(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for region, grp in df.groupby(COL_REGION, sort=False):
        for end_key, label in (("i", "i端"), ("j", "j端")):
            sub = grp[grp[COL_END] == end_key]
            if sub.empty:
                continue
            ids = sorted(sub[COL_ELEMENT].unique())
            rows.append({
                COL_REGION: region,
                COL_END: label,
                "单元个数": len(ids),
                "Midas单元序列": format_midas_sequence(ids),
            })
    return pd.DataFrame(rows)


def prompt_excel_path() -> Path:
    raw = DEFAULT_EXCEL_PATH
    if len(sys.argv) > 1:
        raw = sys.argv[1].strip().strip('"').strip("'")
    elif not raw or not Path(raw).is_file():
        raw = input("请输入单元区域 Excel 路径: ").strip().strip('"').strip("'")
    path = Path(raw).expanduser().resolve()
    if not path.is_file():
        raise SystemExit(f"文件不存在: {path}")
    return path


def main():
    print("=" * 55)
    print("Midas 单元号序列导出（i端 / j端）")
    print("=" * 55)

    excel_path = prompt_excel_path()
    sheet_name = DEFAULT_SHEET_NAME
    print(f"读取: {excel_path}")
    print(f"Sheet: {sheet_name}")

    df = read_element_table(excel_path, sheet_name)
    i_ids, j_ids = split_by_end(df)
    seq_i = format_midas_sequence(i_ids)
    seq_j = format_midas_sequence(j_ids)

    out_dir = excel_path.parent
    path_i = out_dir / OUTPUT_I_FILE
    path_j = out_dir / OUTPUT_J_FILE
    path_detail = out_dir / OUTPUT_DETAIL_FILE

    path_i.write_text(seq_i + "\n", encoding="utf-8")
    path_j.write_text(seq_j + "\n", encoding="utf-8")
    detail = build_region_detail(df)
    detail.to_excel(path_detail, index=False)

    print(f"\n共 {len(df)} 行有效数据")
    print(f"i端: {len(set(i_ids))} 个单元")
    print(f"j端: {len(set(j_ids))} 个单元")

    print("\n" + "-" * 55)
    print("【i端】粘贴到 Midas 单元列:")
    print(seq_i if seq_i else "(无)")
    print("-" * 55)
    print("【j端】粘贴到 Midas 单元列:")
    print(seq_j if seq_j else "(无)")
    print("-" * 55)

    print(f"\n已保存:")
    print(f"  {path_i}")
    print(f"  {path_j}")
    print(f"  {path_detail}（按区域分列）")


if __name__ == "__main__":
    main()
