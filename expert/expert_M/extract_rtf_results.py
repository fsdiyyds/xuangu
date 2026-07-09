"""
Expert N-M 计算书后处理

模式 A（批量）：输入大文件夹 → 遍历其下一级子文件夹
  - 有 list.txt：按 list 排序（匹配时忽略 _ 后缀如 _shear），顶/底→A-A/B-B，Word 标题后追加文件名
  - 无 list.txt：按 RTF 文件名排序，不改 Word 标题
  - 全部子文件夹合并为「{大文件夹名}.docx」存于大文件夹

模式 B（单文件夹）：直接处理单个含 RTF 的文件夹（兼容旧用法）

用法：
  cd expert/expert_M
  python extract_rtf_results.py --folder 计算结果根目录
  python extract_rtf_results.py --folder notes
  python extract_rtf_results.py --folder notes --no-docx
"""

import argparse
import re
import shutil
import sys
import tempfile
import time
import uuid
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple, TypeVar

import pandas as pd

try:
    import pywintypes
except ImportError:
    pywintypes = None

T = TypeVar("T")

RPC_E_CALL_REJECTED = -2147418111
COM_RETRY_COUNT = 60
COM_RETRY_DELAY = 0.2
WORD_QUIT_WAIT = 0.8
COM_RETRY_HRESULTS = frozenset({
    RPC_E_CALL_REJECTED,
    -2147023174,
    -2147023179,
})
DOC_READY_TIMEOUT = 45.0

SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_RTF_FOLDER = SCRIPT_DIR
DEFAULT_OUTPUT_NAME = "rtf_extract_result.xlsx"
LIST_FILE_NAME = "list.txt"

EXCEL_COLUMNS = [
    "ID", "b_cm", "h_cm", "d_cm",
    "As1_cm2", "As2_cm2", "As_min_cm2", "As_max_cm2",
    "rho_pct", "rho_min_pct", "rho_max_pct",
]

WD_FORMAT_XML = 12
WD_COLLAPSE_END = 0
WD_ALERTS_NONE = 0
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3


def _para_plain_text(para) -> str:
    """段落可见文本（去掉 Word 段落结束符等）。"""
    text = para.Range.Text
    for ch in ("\r", "\x07", "\x0b", "\f"):
        text = text.replace(ch, "")
    return text.strip()


def _para_is_empty(para) -> bool:
    return not _para_plain_text(para)


def remove_empty_paragraphs(doc):
    """删除文档中全部空段落（自后向前，至少保留一段）。"""
    def _remove():
        while doc.Paragraphs.Count > 1:
            deleted = False
            for i in range(doc.Paragraphs.Count, 0, -1):
                if doc.Paragraphs.Count <= 1:
                    break
                para = doc.Paragraphs(i)
                if _para_is_empty(para):
                    para.Range.Delete()
                    deleted = True
            if not deleted:
                break

    com_call(_remove, "删除空行")


def trim_trailing_empty_paragraphs(doc):
    """删除文档末尾连续空段落（合并前清理尾部）。"""
    def _trim():
        while doc.Paragraphs.Count > 1:
            last = doc.Paragraphs(doc.Paragraphs.Count)
            if not _para_is_empty(last):
                break
            last.Range.Delete()

    com_call(_trim, "清理末尾空行")


def append_separator_paragraph(doc):
    """在文档末尾插入恰好一个空行（用于合并时分隔各篇）。"""
    end_range = com_call(lambda: doc.Range(), "合并定位")
    com_call(lambda: end_range.Collapse(WD_COLLAPSE_END), "合并折叠")
    com_call(lambda: end_range.InsertParagraphAfter(), "插入分隔空行")


def transform_stem(stem: str) -> str:
    """文件名去后缀后：顶→A-A，底→B-B。"""
    return stem.replace("顶", "A-A").replace("底", "B-B")


def rtf_stem_for_list_match(stem: str) -> str:
    """list.txt 匹配用：忽略第一个 _ 及之后内容（如 C0底 1#_shear → C0底 1#）。"""
    idx = stem.find("_")
    return stem[:idx] if idx >= 0 else stem


class ProgressTracker:
    """终端进度条（单行刷新）。"""

    def __init__(self, total: int, desc: str = "进度", width: int = 36):
        self.total = max(int(total), 1)
        self.current = 0
        self.desc = desc
        self.width = width
        self._active = False

    def update(self, step: int = 1, label: str = ""):
        self.current = min(self.current + step, self.total)
        self._active = True
        pct = self.current / self.total
        filled = int(self.width * pct)
        bar = "█" * filled + "░" * (self.width - filled)
        tail = f"  {label}" if label else ""
        print(
            f"\r  {self.desc} [{bar}] {self.current}/{self.total} "
            f"({pct * 100:.0f}%){tail}",
            end="",
            flush=True,
        )

    def finish(self, message: str = ""):
        if self._active:
            print()
            self._active = False
        if message:
            print(message)


def read_rtf_text(path: Path) -> str:
    return path.read_text(encoding="latin-1", errors="replace")


def _first_float(pattern: str, text: str) -> Optional[float]:
    m = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
    if not m:
        return None
    try:
        return float(m.group(1))
    except ValueError:
        return None


def extract_steel_block(rtf_text: str) -> str:
    m = re.search(
        r"Sections d.+?Acier:(.*?)(?:Analyse par)",
        rtf_text,
        re.IGNORECASE | re.DOTALL,
    )
    return m.group(1) if m else rtf_text


def extract_rtf_fields(rtf_text: str) -> Dict[str, Optional[float]]:
    steel = extract_steel_block(rtf_text)
    return {
        "b_cm": _first_float(r"b\s*=\s*([\d.]+)\s*(?:\\tab\s*)?\(cm\)", rtf_text),
        "h_cm": _first_float(r"h\s*=\s*([\d.]+)\s*(?:\\tab\s*)?\(cm\)", rtf_text),
        "d_cm": _first_float(r"(?<![a-z])d\s*=\s*([\d.]+)\s*(?:\\tab\s*)?\(cm\)", rtf_text),
        "As1_cm2": _first_float(r"dn4\s*s1[^=]*=\s*([\d.]+)\s*\(cm2\)", steel),
        "As2_cm2": _first_float(r"dn4\s*s2[^=]*=\s*([\d.]+)\s*\(cm2\)", steel),
        "As_min_cm2": _first_float(r"s\s*min[^=]*=\s*([\d.]+)\s*\(cm2\)", steel),
        "As_max_cm2": _first_float(r"s\s*max[^=]*=\s*([\d.]+)\s*\(cm2\)", steel),
        "rho_pct": _first_float(r"th\\'e9orique[^=]*=\s*([\d.]+)\s*\(%\)", steel),
        "rho_min_pct": _first_float(r"minimum[^=]*min[^=]*=\s*([\d.]+)\s*\(%\)", steel),
        "rho_max_pct": _first_float(r"maximum[^=]*max[^=]*=\s*([\d.]+)\s*\(%\)", steel),
    }


def list_rtf_files(folder: Path) -> List[Path]:
    return sorted(
        p for p in folder.glob("*.rtf")
        if p.is_file() and not p.name.startswith("~$")
    )


def read_list_order(folder: Path) -> List[str]:
    list_path = folder / LIST_FILE_NAME
    if not list_path.is_file():
        return []
    entries: List[str] = []
    for line in list_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#"):
            entries.append(line)
    return entries


def list_entry_to_regex(entry: str) -> re.Pattern:
    entry = re.sub(r"\s+", " ", entry.strip())
    parts = entry.split(" ")
    if len(parts) >= 2:
        prefix, suffix = parts[0], " ".join(parts[1:])
        pattern = (
            f"^{re.escape(prefix)}"
            f"(?:A-A|B-B|顶|底)?"
            f"\\s+{re.escape(suffix)}$"
        )
        return re.compile(pattern, re.IGNORECASE)
    return re.compile(f"^{re.escape(entry)}$", re.IGNORECASE)


def find_rtf_for_list_entry(
    rtf_by_stem: Dict[str, Path],
    entry: str,
) -> Optional[Path]:
    entry = entry.strip()
    if not entry:
        return None

    transformed = transform_stem(entry)
    for stem, path in rtf_by_stem.items():
        match_key = rtf_stem_for_list_match(stem)
        match_trans = transform_stem(match_key)
        if match_key == entry or match_key == transformed:
            return path
        if match_trans == transformed:
            return path

    regex = list_entry_to_regex(entry)
    matches = [
        stem for stem in rtf_by_stem
        if regex.match(rtf_stem_for_list_match(stem))
        or regex.match(transform_stem(rtf_stem_for_list_match(stem)))
    ]
    if not matches:
        return None
    if len(matches) > 1:
        print(
            f"  警告: list 条目「{entry}」匹配多个 RTF: {matches}，取第一个"
        )
    return rtf_by_stem[matches[0]]


def order_rtf_files(
    rtf_files: List[Path],
    list_entries: List[str],
) -> List[Path]:
    """有 list.txt 则按其中顺序；否则按文件名排序。"""
    if not list_entries:
        return sorted(rtf_files, key=lambda p: p.name.lower())

    rtf_by_stem = {p.stem: p for p in rtf_files}
    used: set = set()
    ordered: List[Path] = []

    for entry in list_entries:
        matched = find_rtf_for_list_entry(rtf_by_stem, entry)
        if matched and matched not in used:
            ordered.append(matched)
            used.add(matched)
            print(f"  list: {entry} → {matched.name}")
        elif matched is None:
            print(f"  [未匹配] list: {entry}")

    for rtf in sorted(rtf_files, key=lambda p: p.name.lower()):
        if rtf not in used:
            ordered.append(rtf)
            print(f"  追加（未在 list 中）: {rtf.name}")

    return ordered


def find_rtf_work_dir(subfolder: Path) -> Optional[Path]:
    """RTF 位于子文件夹根目录或 notes/ 下。"""
    notes = subfolder / "notes"
    if notes.is_dir() and list_rtf_files(notes):
        return notes
    if list_rtf_files(subfolder):
        return subfolder
    return None


def find_list_txt_dirs(subfolder: Path, rtf_work_dir: Path) -> List[Path]:
    """list.txt 可能在子文件夹根目录或 RTF 同目录。"""
    dirs = []
    for d in (subfolder, rtf_work_dir):
        if (d / LIST_FILE_NAME).is_file() and d not in dirs:
            dirs.append(d)
    return dirs


def read_list_order_for_subfolder(subfolder: Path, rtf_work_dir: Path) -> List[str]:
    for d in find_list_txt_dirs(subfolder, rtf_work_dir):
        entries = read_list_order(d)
        if entries:
            return entries
    return []


def subfolder_merged_name(subfolder: Path) -> str:
    return f"{subfolder.name}合并.docx"


def com_call(func: Callable[[], T], label: str = "") -> T:
    last_err: Optional[Exception] = None
    for attempt in range(COM_RETRY_COUNT):
        try:
            return func()
        except Exception as exc:
            last_err = exc
            retryable = (
                pywintypes is not None
                and isinstance(exc, pywintypes.com_error)
                and exc.hresult in COM_RETRY_HRESULTS
            )
            if retryable and attempt < COM_RETRY_COUNT - 1:
                time.sleep(COM_RETRY_DELAY)
                continue
            if label:
                raise RuntimeError(f"Word 操作失败 ({label}): {exc}") from exc
            raise
    if label:
        raise RuntimeError(f"Word 操作超时 ({label}): {last_err}")
    raise last_err  # type: ignore[misc]


def create_word_app():
    import win32com.client
    return win32com.client.Dispatch("Word.Application")


def configure_word_app(word_app):
    word_app.Visible = False
    word_app.DisplayAlerts = WD_ALERTS_NONE
    try:
        word_app.ScreenUpdating = False
    except Exception:
        pass
    try:
        word_app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
    except Exception:
        pass
    try:
        word_app.Options.ConfirmConversions = False
    except Exception:
        pass


def quit_word_app(word_app, hard_kill: bool = True):
    if word_app is None:
        if hard_kill:
            force_kill_word_processes()
        return
    try:
        count = word_app.Documents.Count
        for _ in range(count):
            try:
                word_app.Documents(1).Close(False)
            except Exception:
                break
    except Exception:
        pass
    try:
        word_app.Quit()
    except Exception:
        pass
    time.sleep(WORD_QUIT_WAIT)
    if hard_kill:
        force_kill_word_processes()


def force_kill_word_processes():
    import subprocess
    subprocess.run(
        ["taskkill", "/F", "/IM", "WINWORD.EXE"],
        capture_output=True,
        check=False,
    )
    time.sleep(0.5)


def _doc_full_path(doc) -> Optional[str]:
    try:
        return str(Path(doc.FullName).resolve()).lower()
    except Exception:
        return None


def close_open_document(word_app, path: Path):
    target = str(path.resolve()).lower()
    try:
        count = word_app.Documents.Count
    except Exception:
        return
    for i in range(count, 0, -1):
        try:
            doc = word_app.Documents(i)
            full = _doc_full_path(doc)
            if full and full == target:
                com_call(lambda d=doc: d.Close(False), f"关闭已打开 {path.name}")
        except Exception:
            continue


def remove_file_if_exists(path: Path):
    if not path.is_file():
        return
    try:
        path.unlink()
    except OSError as exc:
        print(f"  警告: 无法删除已存在文件 {path.name}: {exc}")


def _copy_rtf_to_temp(rtf_path: Path) -> Path:
    tmp_dir = Path(tempfile.gettempdir()) / "expert_rtf_convert"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    tmp_rtf = tmp_dir / f"conv_{uuid.uuid4().hex}.rtf"
    shutil.copy2(rtf_path, tmp_rtf)
    return tmp_rtf


def word_open(word_app, path: Path, read_only: bool = False) -> Tuple[Any, Optional[Path]]:
    tmp_rtf: Optional[Path] = None
    open_path = path
    if path.suffix.lower() == ".rtf":
        tmp_rtf = _copy_rtf_to_temp(path)
        open_path = tmp_rtf

    def _open():
        return word_app.Documents.Open(
            FileName=str(open_path.resolve()),
            ConfirmConversions=False,
            ReadOnly=read_only,
            AddToRecentFiles=False,
            Visible=False,
            NoEncodingDialog=True,
            OpenAndRepair=True,
        )

    doc = com_call(_open, f"打开 {path.name}")
    if not wait_document_ready(doc):
        raise RuntimeError(f"Word 打开后文档未就绪: {path.name}")
    return doc, tmp_rtf


def wait_document_ready(doc, timeout: float = DOC_READY_TIMEOUT) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            if doc.Paragraphs.Count > 0:
                time.sleep(0.15)
                return True
        except Exception:
            pass
        time.sleep(0.2)
    return False


def append_title_suffix(doc, suffix: str):
    para = com_call(lambda: doc.Paragraphs(1), "读取标题段")
    rng = para.Range
    text = rng.Text
    suffix_text = f" {suffix}"
    if text.endswith("\r"):
        end_pos = rng.End - 1
        com_call(
            lambda: doc.Range(end_pos, end_pos).InsertAfter(suffix_text),
            "追加标题",
        )
    else:
        com_call(lambda: rng.InsertAfter(suffix_text), "追加标题")


def build_stem_to_id(ordered_rtfs: List[Path], has_list: bool) -> Dict[str, str]:
    """有 list.txt：顶/底→A-A/B-B（忽略 _ 后缀）并重命名 docx；无 list：保持 RTF 原名。"""
    if has_list:
        return {
            rtf.stem: transform_stem(rtf_stem_for_list_match(rtf.stem))
            for rtf in ordered_rtfs
        }
    return {rtf.stem: rtf.stem for rtf in ordered_rtfs}


def convert_rtf_to_named_docx(
    rtf_path: Path,
    output_stem: str,
    word_app,
    append_title: bool = False,
) -> Path:
    out_path = rtf_path.parent / f"{output_stem}.docx"
    close_open_document(word_app, out_path)
    remove_file_if_exists(out_path)

    doc = None
    tmp_rtf: Optional[Path] = None
    try:
        doc, tmp_rtf = word_open(word_app, rtf_path, read_only=False)
        if append_title:
            try:
                append_title_suffix(doc, output_stem)
            except Exception as exc:
                print(f"    [警告] 标题追加失败，继续另存: {exc}", flush=True)
        remove_empty_paragraphs(doc)
        trim_trailing_empty_paragraphs(doc)
        append_separator_paragraph(doc)
        com_call(
            lambda: doc.SaveAs2(
                FileName=str(out_path.resolve()),
                FileFormat=WD_FORMAT_XML,
                AddToRecentFiles=False,
            ),
            f"另存为 {out_path.name}",
        )
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if tmp_rtf is not None and tmp_rtf.is_file():
            try:
                tmp_rtf.unlink()
            except OSError:
                pass

    old_docx = rtf_path.with_suffix(".docx")
    if old_docx.resolve() != out_path.resolve():
        remove_file_if_exists(old_docx)
    return out_path


def convert_rtf_isolated(
    rtf_path: Path,
    output_stem: str,
    append_title: bool = False,
) -> Path:
    word = create_word_app()
    configure_word_app(word)
    try:
        return convert_rtf_to_named_docx(
            rtf_path, output_stem, word, append_title=append_title,
        )
    finally:
        quit_word_app(word, hard_kill=True)


def merge_docx_files(word_app, docx_paths: List[Path], output_path: Path):
    if not docx_paths:
        return

    close_open_document(word_app, output_path)
    remove_file_if_exists(output_path)

    merged = None
    try:
        merged, _ = word_open(word_app, docx_paths[0], read_only=False)
        for docx in docx_paths[1:]:
            end_range = com_call(lambda: merged.Range(), "合并定位")
            com_call(lambda: end_range.Collapse(WD_COLLAPSE_END), "合并折叠")
            com_call(
                lambda p=docx: end_range.InsertFile(
                    FileName=str(p.resolve()),
                    ConfirmConversions=False,
                ),
                f"插入 {docx.name}",
            )
        com_call(
            lambda: merged.SaveAs2(
                FileName=str(output_path.resolve()),
                FileFormat=WD_FORMAT_XML,
                AddToRecentFiles=False,
            ),
            f"保存合并文档 {output_path.name}",
        )
    finally:
        if merged is not None:
            com_call(lambda: merged.Close(False), "关闭合并文档")


def merge_docx_isolated(docx_paths: List[Path], output_path: Path):
    word = create_word_app()
    configure_word_app(word)
    try:
        merge_docx_files(word, docx_paths, output_path)
    finally:
        quit_word_app(word, hard_kill=True)


def warn_duplicate_transform_ids(stem_to_id: Dict[str, str]):
    seen: Dict[str, str] = {}
    for stem, tid in stem_to_id.items():
        if tid in seen:
            print(
                f"  警告: 多个 RTF 将输出同一 docx「{tid}.docx」: "
                f"{seen[tid]} 与 {stem}"
            )
        else:
            seen[tid] = stem


def extract_excel_for_rtfs(
    rtf_files: List[Path],
    stem_to_id: Dict[str, str],
    out_path: Path,
) -> None:
    rows: List[Dict[str, Any]] = []
    progress = ProgressTracker(len(rtf_files), "解析 RTF")
    for rtf in rtf_files:
        note_id = stem_to_id[rtf.stem]
        try:
            fields = extract_rtf_fields(read_rtf_text(rtf))
            rows.append({"ID": note_id, **fields})
        except Exception as exc:
            rows.append({"ID": note_id})
            progress.update(1, f"失败 {rtf.name}")
            continue
        progress.update(1, rtf.name)
    progress.finish(f"  Excel: {out_path}（{len(rows)} 行）")
    df = pd.DataFrame(rows, columns=EXCEL_COLUMNS)
    df.to_excel(out_path, index=False)


def convert_rtfs_to_docx(
    ordered_rtfs: List[Path],
    stem_to_id: Dict[str, str],
    append_title: bool = False,
) -> Tuple[int, int, Dict[str, Path], List[Path]]:
    """RTF→DOCX，返回 docx_map 及按顺序的 docx 列表。"""
    ok, fail = 0, 0
    docx_map: Dict[str, Path] = {}
    ordered_docx: List[Path] = []

    warn_duplicate_transform_ids(stem_to_id)
    if append_title:
        print("  RTF → DOCX（list.txt：顶/底→A-A/B-B，标题追加文件名）...")
    else:
        print("  RTF → DOCX（按文件名排序，不改标题）...")

    progress = ProgressTracker(len(ordered_rtfs), "RTF→DOCX")
    for rtf in ordered_rtfs:
        out_stem = stem_to_id[rtf.stem]
        try:
            out = convert_rtf_isolated(rtf, out_stem, append_title=append_title)
            docx_map[out_stem] = out
            ordered_docx.append(out)
            ok += 1
            progress.update(1, f"{rtf.name} → {out_stem}.docx")
        except Exception as exc:
            progress.update(1, f"失败 {rtf.name}")
            print(f"\n    [失败] {rtf.name}: {exc}")
            force_kill_word_processes()
            fail += 1
    progress.finish(f"  DOCX 转换: 成功 {ok}，失败 {fail}")

    return ok, fail, docx_map, ordered_docx


def process_subfolder(subfolder: Path, save_docx: bool = True) -> Optional[Path]:
    """
    处理单个子文件夹：提取 Excel、RTF→DOCX、合并为「{子文件夹名}合并.docx」。
    返回子文件夹合并 docx 路径（失败或无 RTF 则 None）。
    """
    rtf_work_dir = find_rtf_work_dir(subfolder)
    if rtf_work_dir is None:
        print(f"  跳过「{subfolder.name}」: 无 RTF 文件")
        return None

    rtf_files = list_rtf_files(rtf_work_dir)
    list_entries = read_list_order_for_subfolder(subfolder, rtf_work_dir)

    print(f"\n--- 子文件夹: {subfolder.name} ---")
    print(f"  RTF 目录: {rtf_work_dir}")
    if list_entries:
        print(f"  使用 {LIST_FILE_NAME}（{len(list_entries)} 条）排序，标题追加文件名")
    else:
        print(f"  无 {LIST_FILE_NAME}，按 RTF 文件名排序（不改标题）")

    has_list = bool(list_entries)
    ordered_rtfs = order_rtf_files(rtf_files, list_entries)
    stem_to_id = build_stem_to_id(ordered_rtfs, has_list)

    excel_path = subfolder / DEFAULT_OUTPUT_NAME
    extract_excel_for_rtfs(ordered_rtfs, stem_to_id, excel_path)

    if not save_docx:
        return None

    try:
        import win32com.client  # noqa: F401
    except ImportError:
        print("  警告: 未安装 pywin32，跳过 Word 处理")
        return None

    ok, fail, _, ordered_docx = convert_rtfs_to_docx(
        ordered_rtfs, stem_to_id, append_title=has_list,
    )

    if not ordered_docx:
        print("  警告: 无可用 docx，跳过子文件夹合并")
        return None

    sub_merged = subfolder / subfolder_merged_name(subfolder)
    print(f"  合并 → {sub_merged.name}（{len(ordered_docx)} 篇）...")
    try:
        merge_docx_isolated(ordered_docx, sub_merged)
        print(f"  [OK] {sub_merged}")
        return sub_merged
    except Exception as exc:
        print(f"  [失败] 子文件夹合并: {exc}")
        force_kill_word_processes()
        return None


def list_immediate_subfolders(parent: Path) -> List[Path]:
    return sorted(
        p for p in parent.iterdir()
        if p.is_dir() and not p.name.startswith(".")
    )


def has_processable_subfolders(parent: Path) -> bool:
    """大文件夹下是否存在含 RTF 的直接子文件夹。"""
    return any(find_rtf_work_dir(sub) is not None for sub in list_immediate_subfolders(parent))


def process_parent_folder(parent: Path, save_docx: bool = True) -> None:
    """遍历大文件夹下一级子文件夹，并最终合并总 docx。"""
    if not parent.is_dir():
        raise SystemExit(f"文件夹不存在: {parent}")

    subfolders = list_immediate_subfolders(parent)
    if not subfolders:
        raise SystemExit(f"{parent} 下无子文件夹")

    print(f"\n大文件夹: {parent.name}")
    print(f"子文件夹数: {len(subfolders)}")

    sub_merged_paths: List[Path] = []
    sub_progress = ProgressTracker(len(subfolders), "子文件夹")
    for sub in subfolders:
        merged = process_subfolder(sub, save_docx=save_docx)
        sub_progress.update(1, sub.name)
        if merged and merged.is_file():
            sub_merged_paths.append(merged)
    sub_progress.finish()

    if not save_docx or not sub_merged_paths:
        print("\n未生成子文件夹合并 docx，跳过总合并")
        return

    final_path = parent / f"{parent.name}.docx"
    print(f"\n{'=' * 55}")
    print(f"总合并 → {final_path.name}（{len(sub_merged_paths)} 个子文件夹）")
    for p in sub_merged_paths:
        print(f"  + {p.name}")

    try:
        merge_docx_isolated(sub_merged_paths, final_path)
        print(f"\n[完成] 总文档: {final_path}")
    except Exception as exc:
        print(f"\n[失败] 总合并: {exc}")
        force_kill_word_processes()


def process_single_folder(folder: Path, save_docx: bool = True) -> Path:
    """兼容旧用法：直接处理单个文件夹（无上级批量结构）。"""
    if not folder.is_dir():
        raise SystemExit(f"文件夹不存在: {folder}")

    rtf_work_dir = find_rtf_work_dir(folder) or folder
    rtf_files = list_rtf_files(rtf_work_dir)
    if not rtf_files:
        raise SystemExit(f"在 {folder} 未找到 .rtf 文件")

    list_entries = read_list_order_for_subfolder(folder, rtf_work_dir)
    has_list = bool(list_entries)
    ordered_rtfs = order_rtf_files(rtf_files, list_entries)
    stem_to_id = build_stem_to_id(ordered_rtfs, has_list)

    out_path = folder / DEFAULT_OUTPUT_NAME
    extract_excel_for_rtfs(ordered_rtfs, stem_to_id, out_path)

    if save_docx:
        ok, fail, _, ordered_docx = convert_rtfs_to_docx(
            ordered_rtfs, stem_to_id, append_title=has_list,
        )
        print(f"\nDOCX 完成: 成功 {ok}，失败 {fail}")
        if ordered_docx:
            combined = folder / f"{folder.name}合并.docx"
            merge_docx_isolated(ordered_docx, combined)
            print(f"合并完成: {combined}")

    return out_path


def parse_args():
    parser = argparse.ArgumentParser(
        description="Expert RTF 批量后处理：子文件夹 RTF→DOCX→合并"
    )
    parser.add_argument(
        "--folder",
        type=str,
        default=str(DEFAULT_RTF_FOLDER),
        help="大文件夹路径（遍历其下一级子文件夹；若无子文件夹则处理自身）",
    )
    parser.add_argument(
        "--no-docx",
        action="store_true",
        help="仅提取 Excel，不处理 Word",
    )
    return parser.parse_args()


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass

    args = parse_args()
    folder = Path(args.folder).expanduser().resolve()
    save_docx = not args.no_docx

    print("=" * 55)
    print("Expert RTF 结果提取 / Word 合并")
    print("=" * 55)
    print(f"输入: {folder}")

    if has_processable_subfolders(folder):
        process_parent_folder(folder, save_docx=save_docx)
    else:
        print("（单文件夹模式）")
        process_single_folder(folder, save_docx=save_docx)


if __name__ == "__main__":
    main()
