"""
桥面板内力提取 → Word 表格

任务 1：弯矩表（Mxx 纵向取各横向位置最大值，Myy 分列）
任务 2：剪力表（Vxx 纵向合并，Vyy 分列）
任务 3：生成 expert_M 用 input_NM_纵向.xlsx / input_NM_横向.xlsx
        （每 ID：ELU/ELS 各 max+min 共 4 组；横向 × 表个数；N=0）

用法：
  cd expert/internal_force_extract
  python main.py 输入.xlsx
  python main.py 输入.xlsx --spacing 1.0
  python main.py 输入.xlsx --mapping load_mapping.txt
  python main.py 输入.xlsx --no-input-nm   # 仅 Word，不导出 input_NM
"""
import argparse
import sys
from pathlib import Path

import pandas as pd

SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_MAPPING = SCRIPT_DIR / "load_mapping.txt"


def main():
    parser = argparse.ArgumentParser(description="Midas 板单元内力 → Word 表格")
    parser.add_argument("excel", nargs="?", help="含「跨度信息」「内力提取*」的 Excel")
    parser.add_argument(
        "--spacing", type=float, default=1.0,
        help="普通桥跨纵桥向提取间距 m（默认 1.0）",
    )
    parser.add_argument(
        "--mapping", type=Path, default=DEFAULT_MAPPING,
        help="荷载映射文件 load_mapping.txt",
    )
    parser.add_argument(
        "--output-dir", type=Path, default=SCRIPT_DIR / "output",
        help="Word 输出目录",
    )
    parser.add_argument(
        "--create-sample", action="store_true",
        help="生成示例 Excel 后退出",
    )
    parser.add_argument(
        "--no-input-nm", action="store_true",
        help="不生成 input_NM.xlsx",
    )
    args = parser.parse_args()

    if args.create_sample:
        from create_sample import write_sample
        out = SCRIPT_DIR / "sample_input.xlsx"
        write_sample(out)
        print(f"已生成示例: {out}")
        return

    excel = Path(args.excel) if args.excel else SCRIPT_DIR / "sample_input.xlsx"
    if not excel.is_file():
        print(f"未找到输入文件: {excel}")
        print("可运行: python main.py --create-sample")
        sys.exit(1)

    mapping = args.mapping
    if not mapping.is_file():
        print(f"未找到荷载映射: {mapping}")
        sys.exit(1)

    sys.path.insert(0, str(SCRIPT_DIR))
    from aggregate import build_force_tables
    from word_export import build_moment_document, build_shear_document, save_document

    out_dir = args.output_dir
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"读取: {excel}")
    print(f"荷载映射: {mapping}")
    print(f"桥跨间距: {args.spacing} m")

    _, moment_rows, n_trans = build_force_tables(
        excel, mapping, normal_spacing=args.spacing, mode="moment",
    )
    print(f"  → 弯矩：{len(moment_rows)} 个纵桥向站点，{n_trans} 个横向位置")
    mom_doc = build_moment_document(moment_rows, n_trans)
    mom_path = out_dir / "efforts_moments.docx"
    save_document(mom_doc, mom_path)
    print(f"  → 已保存: {mom_path}")

    _, shear_rows, n_trans2 = build_force_tables(
        excel, mapping, normal_spacing=args.spacing, mode="shear",
    )
    shear_doc = build_shear_document(shear_rows, n_trans2)
    shear_path = out_dir / "efforts_cisaillements.docx"
    save_document(shear_doc, shear_path)
    print(f"  → 已保存: {shear_path}")

    if not args.no_input_nm:
        from export_input_nm import export_input_nm
        paths = export_input_nm(excel, mapping, out_dir, normal_spacing=args.spacing)
        long_df = pd.read_excel(paths["longitudinal"])
        trans_df = pd.read_excel(paths["transverse"])
        n_ids = long_df["ID"].nunique()
        print(f"  → 已保存: {paths['longitudinal']}（{len(long_df)} 行，{n_ids} 个 ID × 4 组纵向荷载）")
        print(f"  → 已保存: {paths['transverse']}（{len(trans_df)} 行，{n_ids} 个 ID × {len(trans_df)//max(n_ids,1)} 组横向荷载）")

    print("完成。")


if __name__ == "__main__":
    main()
