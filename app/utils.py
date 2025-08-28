import re
from typing import Iterable, List, Tuple

def normalize_ws(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()

def sliding_window(tokens: List[str], size: int, overlap: int) -> Iterable[Tuple[int, int]]:
    step = max(1, size - overlap)
    for start in range(0, max(0, len(tokens)-1), step):
        end = min(len(tokens), start + size)
        yield start, end
        if end == len(tokens): break

def md_title_path(path_parts: List[str]) -> str:
    return " / ".join([p.strip() for p in path_parts if p.strip()])
