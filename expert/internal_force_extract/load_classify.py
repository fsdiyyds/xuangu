"""荷载名称 → ULS/ALS/SLS，并识别包络最大/最小。"""
import re
from pathlib import Path
from typing import Dict, Optional, Tuple

LOAD_TYPES = frozenset({"ULS", "ALS", "SLS"})

# Word 表格组合行
WORD_COMBO = {"ULS": "ELU", "ALS": "ELU", "SLS": "ELS"}


def read_load_mapping(path: Path) -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    text = path.read_text(encoding="utf-8")
    for line in text.splitlines():
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
            raise ValueError(f"未知荷载类型 '{cat}'，仅支持 {sorted(LOAD_TYPES)}")
        mapping[name] = cat
    return mapping


def _norm_key(name: str) -> str:
    return re.sub(r"\s+", "", str(name).strip()).upper()


def classify_load(
    load_name: str, mapping: Dict[str, str],
) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """
    返回 (ULS|ALS|SLS, MAX|MIN|None, Word组合 ELU|ELS|None)。
    MAX/MIN 由荷载名中「最大」「最小」或 (最大)/(最小) 判定。
    """
    name = str(load_name).strip()
    if not name:
        return None, None, None

    envelope = None
    if "最大" in name or "(MAX" in name.upper():
        envelope = "MAX"
    elif "最小" in name or "(MIN" in name.upper():
        envelope = "MIN"

    nkey = _norm_key(name)
    cat = None
    for key, c in sorted(mapping.items(), key=lambda x: len(x[0]), reverse=True):
        kn = _norm_key(key)
        if name == key or nkey == kn or nkey.startswith(kn) or kn in nkey:
            cat = c
            break

    if cat is None:
        upper = name.upper()
        if "ELUA" in upper or "地震" in name:
            cat = "ALS"
        elif "ELU" in upper or "ULS" in upper:
            cat = "ULS"
        elif "ELS" in upper or "SLS" in upper or "准永久" in name:
            cat = "SLS"

    word = WORD_COMBO.get(cat) if cat else None
    return cat, envelope, word
