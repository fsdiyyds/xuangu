"""
Midas 单元内力归集

- 遍历文件夹内所有 .xlsx
- 自动识别每个文件中的全部 sheet，同名 sheet 跨文件合并
- 荷载经 load_mapping.txt 映射为 ULS / ALS / SLS
- 同一单元：在各类别全部行中，对 N/Qz/Qy/My/Mz 各列分别取 MAX/MIN 及并发内力
- 不按「成分」列筛选；输出表头：编号、名称、类型、N、Qz、Qy、My、Mz
"""

import os
import re
from collections import defaultdict
from pathlib import Path
from typing import DefaultDict, Dict, List, Optional, Set, Tuple, Union

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, PatternFill, Side

PathLike = Union[str, Path]

# ===================== 运行配置 =====================
DEFAULT_INPUT_FOLDER = ""
# 与 Excel 放在同一文件夹内
MAPPING_FILE_NAME = "load_mapping.txt"
OUTPUT_FILE_NAME = "force_aggregate_result.xlsx"
# 结果文件中排版 sheet 后缀（合并单元格，便于粘贴 Word）
REPORT_SHEET_SUFFIX = "_排版"
# 源 Excel 极值行填充色：MAX 浅黄，MIN 浅蓝；同一行兼有两种极值时用浅橙
MARK_FILL_MAX = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
MARK_FILL_MIN = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
MARK_FILL_BOTH = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
# ==================================================

COL_ELEMENT = "单元"
COL_LOAD = "荷载"
COL_POSITION = "位置"
COL_COMPONENT = "成分"

FORCE_COLS = [
    "轴向 (kN)",
    "剪力-y (kN)",
    "剪力-z (kN)",
    "扭矩 (kN·m)",
    "弯矩-y (kN·m)",
    "弯矩-z (kN·m)",
]

REQUIRED_COLS = [COL_ELEMENT, COL_LOAD, COL_POSITION, COL_COMPONENT] + FORCE_COLS

COMPONENT_TO_COL = {
    "轴向": "轴向 (kN)",
    "剪力-y": "剪力-y (kN)",
    "剪力-z": "剪力-z (kN)",
    "扭矩": "扭矩 (kN·m)",
    "弯矩-y": "弯矩-y (kN·m)",
    "弯矩-z": "弯矩-z (kN·m)",
}

# 输出类型 → 源表内力列（在类别全部行中直接比该列极值，不看成分列）
OUTPUT_TYPE_ROWS: List[Tuple[str, str]] = [
    ("N", "轴向 (kN)"),
    ("Qy", "剪力-y (kN)"),
    ("Qz", "剪力-z (kN)"),
    ("My", "弯矩-y (kN·m)"),
    ("Mz", "弯矩-z (kN·m)"),
]

# 输出列（与目标表格一致，顺序固定）
OUTPUT_COLUMNS = [
    "编号",
    "名称",
    "类型",
    "N(kN)",
    "Qz(kNm)",
    "Qy(kN·m)",
    "My(kNm)",
    "Mz(kNm)",
]

COL_ALIASES = {
    "单元": COL_ELEMENT,
    "荷载": COL_LOAD,
    "位置": COL_POSITION,
    "成分": COL_COMPONENT,
    "轴向": "轴向 (kN)",
    "轴向力": "轴向 (kN)",
    "剪力y": "剪力-y (kN)",
    "剪力-y": "剪力-y (kN)",
    "剪力z": "剪力-z (kN)",
    "剪力-z": "剪力-z (kN)",
    "扭矩": "扭矩 (kN·m)",
    "弯矩y": "弯矩-y (kN·m)",
    "弯矩-y": "弯矩-y (kN·m)",
    "弯矩z": "弯矩-z (kN·m)",
    "弯矩-z": "弯矩-z (kN·m)",
}

GROUP_KEYS = [COL_ELEMENT, COL_POSITION]
LOAD_TYPES = ("ULS", "ALS", "SLS")
EXTREMES = ("MAX", "MIN")

# 报告排版：名称块顺序、类型行顺序
REPORT_NAME_ORDER = [f"{lt} {ex}" for lt in LOAD_TYPES for ex in EXTREMES]
REPORT_TYPE_ORDER = [t for t, _ in OUTPUT_TYPE_ROWS]

REPORT_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
REPORT_THIN = Side(style="thin")
REPORT_BORDER = Border(
    left=REPORT_THIN, right=REPORT_THIN, top=REPORT_THIN, bottom=REPORT_THIN
)


def normalize_text(s) -> str:
    if pd.isna(s):
        return ""
    t = str(s).strip()
    t = t.replace("－", "-").replace("—", "-").replace("·", "")
    return re.sub(r"\s+", "", t)


def normalize_component(name: str) -> str:
    raw = str(name).strip().replace("－", "-").replace("—", "-")
    raw_lower = raw.lower()
    short_map = {
        "n": "轴向",
        "qy": "剪力-y",
        "qz": "剪力-z",
        "my": "弯矩-y",
        "mz": "弯矩-z",
        "mt": "扭矩",
        "t": "扭矩",
    }
    if raw_lower in short_map:
        return short_map[raw_lower]
    s = normalize_text(name)
    if s.lower() in short_map:
        return short_map[s.lower()]
    aliases = {
        "剪力y": "剪力-y",
        "剪力z": "剪力-z",
        "弯矩y": "弯矩-y",
        "弯矩z": "弯矩-z",
        "轴力": "轴向",
    }
    return aliases.get(s, raw)


def _norm_col_key(name) -> str:
    return normalize_text(name).lower()


def find_header_row(raw: pd.DataFrame) -> int:
    max_scan = min(len(raw), 30)
    for i in range(max_scan):
        cells = {_norm_col_key(v) for v in raw.iloc[i].tolist()}
        if "单元" in cells and ("荷载" in cells or "位置" in cells):
            return i
    raise ValueError("未识别到表头行（需包含「单元」「荷载」或「位置」列）")


def align_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename = {}
    for col in df.columns:
        key = _norm_col_key(col)
        for alias, std in COL_ALIASES.items():
            if key == _norm_col_key(alias):
                rename[col] = std
                break
        else:
            for alias, std in COL_ALIASES.items():
                if key.startswith(_norm_col_key(alias)):
                    rename[col] = std
                    break
    return df.rename(columns=rename)


def drop_empty_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.dropna(how="all")
    mask_blank = df.apply(
        lambda row: all(normalize_text(v) == "" for v in row),
        axis=1,
    )
    df = df.loc[~mask_blank].copy()
    if COL_ELEMENT not in df.columns:
        return df
    elem = df[COL_ELEMENT].apply(normalize_text)
    df = df.loc[elem != ""].copy()
    df = df.loc[df[COL_ELEMENT].notna()].copy()
    return df


def read_raw_sheet(excel_path: PathLike, sheet_name: str) -> pd.DataFrame:
    raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, dtype=object)
    header_idx = find_header_row(raw)
    header = raw.iloc[header_idx].tolist()
    df = raw.iloc[header_idx + 1:].copy()
    df.columns = header
    # Excel 行号（1-based），用于回写高亮
    df["_source_row"] = [header_idx + 2 + i for i in range(len(df))]
    df = align_columns(df)
    return drop_empty_rows(df)


def read_load_mapping(path: PathLike) -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "," in line:
                name, cat = line.split(",", 1)
            elif "=" in line:
                name, cat = line.split("=", 1)
            else:
                parts = re.split(r"\s+", line, maxsplit=1)
                if len(parts) != 2:
                    continue
                name, cat = parts
            name, cat = name.strip(), cat.strip().upper()
            if cat not in LOAD_TYPES:
                raise ValueError(f"未知荷载类型 '{cat}'，仅支持 {LOAD_TYPES}")
            mapping[name] = cat
    return mapping


def find_mapping_file(folder: Path) -> Path:
    p = folder / MAPPING_FILE_NAME
    if p.is_file():
        return p
    for name in ("荷载映射.txt", "load_map.txt"):
        alt = folder / name
        if alt.is_file():
            return alt
    create_mapping_template(folder / MAPPING_FILE_NAME)
    raise SystemExit(f"请编辑 {folder / MAPPING_FILE_NAME} 后重新运行")


def list_excel_files(folder: Path) -> List[Path]:
    files = []
    for p in sorted(folder.glob("*.xlsx")):
        if p.name.startswith("~$"):
            continue
        if p.name == OUTPUT_FILE_NAME:
            continue
        files.append(p)
    return files


def discover_sheet_names(excel_paths: List[Path]) -> List[str]:
    names: Set[str] = set()
    for p in excel_paths:
        try:
            names.update(pd.ExcelFile(p).sheet_names)
        except Exception as e:
            print(f"  警告：无法打开 {p.name} 读取 sheet 列表: {e}")
    return sorted(names)


def read_midas_sheet(excel_path: PathLike, sheet_name: str) -> pd.DataFrame:
    df = read_raw_sheet(excel_path, sheet_name)
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"缺少列: {missing}")
    df = df.copy()
    df[COL_ELEMENT] = df[COL_ELEMENT].astype(str).str.strip()
    df[COL_LOAD] = df[COL_LOAD].astype(str).str.strip()
    df[COL_POSITION] = df[COL_POSITION].astype(str).str.strip()
    df[COL_COMPONENT] = df[COL_COMPONENT].apply(normalize_component)
    for c in FORCE_COLS:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["_source_file"] = os.path.basename(str(excel_path))
    df["_source_path"] = str(Path(excel_path).resolve())
    df["_sheet_name"] = sheet_name
    return df


def collect_sheet_from_excels(
    excel_paths: List[Path],
    sheet_name: str,
    mapping: Dict[str, str],
) -> Tuple[pd.DataFrame, List[str], Set[str]]:
    parts: List[pd.DataFrame] = []
    errors: List[str] = []
    unmapped: Set[str] = set()

    for p in excel_paths:
        if sheet_name not in pd.ExcelFile(p).sheet_names:
            continue
        try:
            df = read_midas_sheet(p, sheet_name)
        except Exception as e:
            errors.append(f"{p.name}: {e}")
            continue
        mapped = df[COL_LOAD].map(mapping)
        unmapped.update(df.loc[mapped.isna(), COL_LOAD].unique())
        df = df.loc[mapped.notna()].copy()
        df["荷载类型"] = mapped[mapped.notna()].values
        parts.append(df)
        print(f"    {p.name}: {len(df)} 行")

    if not parts:
        return pd.DataFrame(), errors, unmapped

    return pd.concat(parts, ignore_index=True), errors, unmapped


def _concurrent_forces(row: pd.Series) -> Dict[str, float]:
    return {
        "N(kN)": row["轴向 (kN)"],
        "Qz(kNm)": row["剪力-z (kN)"],
        "Qy(kN·m)": row["剪力-y (kN)"],
        "My(kNm)": row["弯矩-y (kN·m)"],
        "Mz(kNm)": row["弯矩-z (kN·m)"],
    }


def _empty_output_row(code: str, name_label: str, type_label: str) -> dict:
    row = {"编号": code, "名称": name_label, "类型": type_label}
    for c in OUTPUT_COLUMNS[3:]:
        row[c] = None
    return row


def _select_extreme_row(
    pool: pd.DataFrame,
    driver_col: str,
    extreme: str,
) -> Optional[pd.Series]:
    """
    在类别内全部行中，取 driver_col 的极值所在行（含并发内力）。
    极值并列时取第一行，保证有数据就不留空。
    """
    if pool.empty:
        return None
    valid = pool.dropna(subset=[driver_col])
    if valid.empty:
        return None
    if extreme == "MAX":
        target = valid[driver_col].max()
        picks = valid.loc[valid[driver_col] == target]
    else:
        target = valid[driver_col].min()
        picks = valid.loc[valid[driver_col] == target]
    return picks.iloc[0]


def _mark_source_row(
    row: pd.Series,
    marked_rows: DefaultDict[str, DefaultDict[str, Dict[int, Set[str]]]],
    extreme: str,
):
    path = row.get("_source_path")
    sheet = row.get("_sheet_name")
    row_num = row.get("_source_row")
    if path and sheet and pd.notna(row_num):
        marked_rows[path][sheet][int(row_num)].add(extreme)


def build_element_summary(
    df: pd.DataFrame,
    element: str,
    position: str,
    marked_rows: DefaultDict[str, DefaultDict[str, Dict[int, Set[str]]]],
) -> List[dict]:
    """
    单个单元：先按映射归入 ULS/ALS/SLS，再在各类别全部行中对
    N/Qz/Qy/My/Mz 各列分别取 MAX/MIN → 30 行输出。
    """
    rows: List[dict] = []
    code = element
    if position and normalize_text(position):
        code = f"{element} {position}"

    sub_all = df[
        (df[COL_ELEMENT] == element) & (df[COL_POSITION] == position)
    ]

    for load_type in LOAD_TYPES:
        sub_cat = sub_all[sub_all["荷载类型"] == load_type]
        for extreme in EXTREMES:
            name_label = f"{load_type} {extreme}"
            if sub_cat.empty:
                for type_label, _ in OUTPUT_TYPE_ROWS:
                    rows.append(_empty_output_row(code, name_label, type_label))
                continue

            for type_label, driver_col in OUTPUT_TYPE_ROWS:
                r = _select_extreme_row(sub_cat, driver_col, extreme)
                if r is None:
                    rows.append(_empty_output_row(code, name_label, type_label))
                    continue

                _mark_source_row(r, marked_rows, extreme)
                out = {
                    "编号": code,
                    "名称": name_label,
                    "类型": type_label,
                }
                out.update(_concurrent_forces(r))
                rows.append(out)

    return rows


def build_summary_table(
    df: pd.DataFrame,
    marked_rows: DefaultDict[str, DefaultDict[str, Dict[int, Set[str]]]],
) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    all_rows: List[dict] = []
    groups = df.groupby(GROUP_KEYS, sort=False)

    for element, position in groups.groups.keys():
        all_rows.extend(
            build_element_summary(df, element, position, marked_rows)
        )

    return pd.DataFrame(all_rows, columns=OUTPUT_COLUMNS)


def apply_source_marks(
    marked_rows: DefaultDict[str, DefaultDict[str, Dict[int, Set[str]]]],
):
    """在源 Excel 中将归集用到的 MAX/MIN 行填充为不同浅色背景"""
    for path_str, sheets in marked_rows.items():
        path = Path(path_str)
        if not path.is_file():
            print(f"  警告：源文件不存在，跳过标记 {path}")
            continue
        wb = load_workbook(path)
        n_max, n_min, n_both = 0, 0, 0
        for sheet_name, row_map in sheets.items():
            if sheet_name not in wb.sheetnames:
                continue
            ws = wb[sheet_name]
            max_col = max(ws.max_column or 1, 1)
            for row_num in sorted(row_map):
                kinds = row_map[row_num]
                if "MAX" in kinds and "MIN" in kinds:
                    fill = MARK_FILL_BOTH
                    n_both += 1
                elif "MAX" in kinds:
                    fill = MARK_FILL_MAX
                    n_max += 1
                else:
                    fill = MARK_FILL_MIN
                    n_min += 1
                for col in range(1, max_col + 1):
                    ws.cell(row=row_num, column=col).fill = fill
        wb.save(path)
        print(
            f"  已标记: {path.name}"
            f"（MAX {n_max} 行 / MIN {n_min} 行 / 兼有 {n_both} 行）"
        )


def sort_report_table(df: pd.DataFrame) -> pd.DataFrame:
    """按编号 → ULS MAX/MIN… → N/Qy/Qz/My/Mz 排序，数值保留两位小数"""
    if df.empty:
        return df.copy()
    name_rank = {n: i for i, n in enumerate(REPORT_NAME_ORDER)}
    type_rank = {t: i for i, t in enumerate(REPORT_TYPE_ORDER)}
    out = df.copy()
    out["_name_rank"] = out["名称"].map(name_rank).fillna(999)
    out["_type_rank"] = out["类型"].map(type_rank).fillna(999)
    out = out.sort_values(
        ["编号", "_name_rank", "_type_rank"],
        kind="stable",
    ).drop(columns=["_name_rank", "_type_rank"])
    for col in OUTPUT_COLUMNS[3:]:
        out[col] = pd.to_numeric(out[col], errors="coerce").round(2)
    return out.reset_index(drop=True)


def _merge_vertical(ws, col: int, start_row: int, end_row: int):
    if end_row > start_row:
        ws.merge_cells(
            start_row=start_row,
            start_column=col,
            end_row=end_row,
            end_column=col,
        )
    ws.cell(start_row, col).alignment = REPORT_CENTER


def apply_report_sheet_format(ws):
    """
    报告排版：编号列按单元合并，名称列按 ULS MAX 等块合并，加边框居中。
    数据从第 2 行开始（第 1 行为表头）。
    """
    max_row = ws.max_row or 1
    max_col = len(OUTPUT_COLUMNS)
    if max_row < 2:
        return

    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(r, c)
            cell.border = REPORT_BORDER
            cell.alignment = REPORT_CENTER

    # 列宽（便于 Word 粘贴）
    widths = {"A": 20, "B": 12, "C": 8, "D": 12, "E": 12, "F": 12, "G": 12, "H": 12}
    for letter, w in widths.items():
        ws.column_dimensions[letter].width = w

    data_start = 2
    row = data_start
    while row <= max_row:
        code = ws.cell(row, 1).value
        code_start = row
        while row <= max_row and ws.cell(row, 1).value == code:
            row += 1
        code_end = row - 1
        _merge_vertical(ws, 1, code_start, code_end)

        name_row = code_start
        while name_row <= code_end:
            name = ws.cell(name_row, 2).value
            name_start = name_row
            while name_row <= code_end and ws.cell(name_row, 2).value == name:
                name_row += 1
            name_end = name_row - 1
            _merge_vertical(ws, 2, name_start, name_end)


def export_all_sheets(sheet_tables: Dict[str, pd.DataFrame], out_path: Path):
    used_names: Set[str] = set()
    report_sheet_names: List[str] = []

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for sheet_name, table in sheet_tables.items():
            safe = sanitize_sheet_name(sheet_name, used_names)
            table.to_excel(writer, sheet_name=safe, index=False)

            report_df = sort_report_table(table)
            report_name = sanitize_sheet_name(
                f"{sheet_name}{REPORT_SHEET_SUFFIX}", used_names
            )
            report_df.to_excel(writer, sheet_name=report_name, index=False)
            report_sheet_names.append(report_name)

    wb = load_workbook(out_path)
    for name in report_sheet_names:
        if name in wb.sheetnames:
            apply_report_sheet_format(wb[name])
    wb.save(out_path)


def sanitize_sheet_name(name: str, used: Set[str]) -> str:
    """Excel sheet 名最长 31 字符，且不可重复"""
    invalid = r'[\\/*?:\[\]]'
    s = re.sub(invalid, "_", str(name)).strip() or "Sheet"
    if len(s) > 31:
        s = s[:31]
    base, n = s, 1
    while s in used:
        suffix = f"_{n}"
        s = (base[:31 - len(suffix)] + suffix) if len(base) + len(suffix) > 31 else base + suffix
        n += 1
    used.add(s)
    return s


def prompt_folder() -> Path:
    raw = DEFAULT_INPUT_FOLDER
    if not raw:
        raw = input("请输入 Excel 文件夹路径: ").strip().strip('"').strip("'")
    folder = Path(raw).expanduser().resolve()
    if not folder.is_dir():
        raise SystemExit(f"文件夹不存在: {folder}")
    return folder


def create_mapping_template(path: Path):
    template = """# 荷载名称 → ULS / ALS / SLS（与 Excel「荷载」列完全一致）
# 格式：荷载名称,类型

ELU包络(最大),ULS
基本组合1,ULS
地震组合E1,ALS
准永久组合,SLS
"""
    path.write_text(template, encoding="utf-8")
    print(f"已生成映射模板: {path}，请按实际荷载修改后重新运行")


def main():
    print("=" * 55)
    print("Midas 单元内力归集（自动遍历全部 sheet）")
    print("=" * 55)

    input_folder = prompt_folder()
    mapping_path = find_mapping_file(input_folder)
    mapping = read_load_mapping(mapping_path)

    excel_paths = list_excel_files(input_folder)
    if not excel_paths:
        raise SystemExit(f"在 {input_folder} 未找到 .xlsx 文件")

    sheet_names = discover_sheet_names(excel_paths)
    if not sheet_names:
        raise SystemExit("未找到任何 sheet")

    print(f"文件夹: {input_folder}")
    print(f"映射文件: {mapping_path.name}（{len(mapping)} 条荷载）")
    print(f"Excel 文件: {len(excel_paths)} 个")
    print(f"发现 sheet: {len(sheet_names)} 个")

    all_unmapped: Set[str] = set()
    sheet_tables: Dict[str, pd.DataFrame] = {}
    marked_rows: DefaultDict[str, DefaultDict[str, Dict[int, Set[str]]]] = defaultdict(
        lambda: defaultdict(lambda: defaultdict(set))
    )

    for sheet_name in sheet_names:
        print(f"\n处理 sheet: {sheet_name}")
        merged, errors, unmapped = collect_sheet_from_excels(
            excel_paths, sheet_name, mapping
        )
        all_unmapped.update(unmapped)

        if errors:
            for msg in errors:
                print(f"    跳过: {msg}")

        if merged.empty:
            print("    无有效数据，跳过")
            continue

        table = build_summary_table(merged, marked_rows)
        sheet_tables[sheet_name] = table
        n_elem = merged.groupby(GROUP_KEYS).ngroups
        print(f"    归集完成: {n_elem} 个单元 → {len(table)} 行输出")

    if all_unmapped:
        print(f"\n警告：以下荷载未在映射表中，已跳过: {sorted(all_unmapped)}")

    if not sheet_tables:
        raise SystemExit("没有生成任何归集结果，请检查数据和映射表")

    out_path = input_folder / OUTPUT_FILE_NAME
    export_all_sheets(sheet_tables, out_path)
    print(f"\n结果表 → {out_path}")
    print(
        f"共 {len(sheet_tables)} 个计算项；"
        f"每个计算项含数据 sheet + {REPORT_SHEET_SUFFIX} 排版 sheet（合并单元格）"
    )

    if marked_rows:
        print("\n标记源 Excel 极值行（MAX 浅黄 / MIN 浅蓝 / 兼有 浅橙）...")
        apply_source_marks(marked_rows)
    else:
        print("\n未找到可标记的极值行")


if __name__ == "__main__":
    main()
