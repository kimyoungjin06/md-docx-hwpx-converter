#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, Iterator, List, Optional, Sequence, Tuple


PANDOC_API_VERSION = [1, 23, 1]


def _run(cmd: Sequence[str], *, cwd: Optional[Path] = None) -> None:
    subprocess.run(list(cmd), check=True, cwd=str(cwd) if cwd else None)


def _ensure_pandoc() -> str:
    pandoc = shutil.which("pandoc")
    if not pandoc:
        raise RuntimeError("pandoc is required but was not found in PATH.")
    return pandoc


def _load_json(path: Path) -> Dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def _write_json(path: Path, data: Dict[str, Any]) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")


def _inlines_to_text(inlines: Sequence[Dict[str, Any]]) -> str:
    parts: List[str] = []
    for el in inlines:
        t = el.get("t")
        if t == "Str":
            parts.append(el.get("c", ""))
        elif t in {"Space", "SoftBreak", "LineBreak"}:
            parts.append(" ")
        elif t in {"Emph", "Strong", "SmallCaps", "Strikeout", "Superscript", "Subscript"}:
            parts.append(_inlines_to_text(el.get("c", [])))
        elif t == "Span":
            # ["Attr", [Inlines]]
            c = el.get("c", [None, []])
            parts.append(_inlines_to_text(c[1] if isinstance(c, list) and len(c) >= 2 else []))
        elif t == "Code":
            parts.append(el.get("c", ["", ""])[1])
        elif t == "Link":
            # ["Attr", [Inlines], [Target]]
            parts.append(_inlines_to_text(el.get("c", [None, [], None])[1]))
        elif t == "Image":
            # ["Attr", [Inlines], [Target]]
            parts.append(_inlines_to_text(el.get("c", [None, [], None])[1]))
        else:
            parts.append(" ")
    return "".join(parts)


def _normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text.replace("\u00a0", " ")).strip()


def _str_para(text: str) -> Dict[str, Any]:
    words = re.split(r"(\s+)", text.strip())
    inlines: List[Dict[str, Any]] = []
    for w in words:
        if not w:
            continue
        if w.isspace():
            if inlines and inlines[-1].get("t") == "Space":
                continue
            inlines.append({"t": "Space"})
            continue
        inlines.append({"t": "Str", "c": w})
    return {"t": "Para", "c": inlines}


def _header_text(block: Dict[str, Any]) -> Optional[str]:
    if block.get("t") != "Header":
        return None
    level, _attr, inlines = block.get("c", [None, None, None])
    if not isinstance(level, int) or not isinstance(inlines, list):
        return None
    return _normalize_text(_inlines_to_text(inlines))


def _find_header_index(
    blocks: Sequence[Dict[str, Any]],
    *,
    needle: str,
    level: Optional[int] = None,
    contains: bool = False,
) -> int:
    needle_norm = _normalize_text(needle)
    for idx, block in enumerate(blocks):
        text = _header_text(block)
        if text is None:
            continue
        if level is not None:
            if block.get("c", [None])[0] != level:
                continue
        if contains:
            if needle_norm in text:
                return idx
        else:
            if needle_norm == text:
                return idx
    raise ValueError(f"Header not found: {needle!r}")


def _section_end(blocks: Sequence[Dict[str, Any]], header_idx: int) -> int:
    header = blocks[header_idx]
    if header.get("t") != "Header":
        raise ValueError("section start must be a Header block")
    level = header.get("c", [None])[0]
    if not isinstance(level, int):
        raise ValueError("Invalid Header block (missing level)")
    for i in range(header_idx + 1, len(blocks)):
        blk = blocks[i]
        if blk.get("t") != "Header":
            continue
        next_level = blk.get("c", [None])[0]
        if isinstance(next_level, int) and next_level <= level:
            return i
    return len(blocks)


def _extract_section_by_header(
    blocks: Sequence[Dict[str, Any]],
    *,
    header_text: str,
    level: Optional[int] = None,
    contains: bool = False,
) -> List[Dict[str, Any]]:
    start = _find_header_index(blocks, needle=header_text, level=level, contains=contains)
    end = _section_end(blocks, start)
    return [json.loads(json.dumps(b)) for b in blocks[start:end]]


def _shift_heading_levels(blocks: Iterable[Dict[str, Any]], delta: int) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    for b in blocks:
        if b.get("t") == "Header":
            level, attr, inlines = b.get("c", [None, None, None])
            if isinstance(level, int):
                level = max(1, min(6, level + delta))
                b = {"t": "Header", "c": [level, attr, inlines]}
        out.append(b)
    return out


_DATABOOK_HEADING_NUM_PREFIX_RE = re.compile(r"^\s*\d+(?:\.\d+)+\s*[.)]?\s+")
_DATABOOK_HEADING_STEP_PREFIX_RE = re.compile(r"^\s*\d+\.\s+")


def _strip_databook_heading_numbers(blocks: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    데이터북 DOCX에서 들어온 Heading 텍스트에는 '2.4.1 1. ...' 같은
    "문자열 번호"가 포함되어 있다. 최종보고서(DOCX)의 Heading 스타일과 충돌하므로
    숫자 prefix는 제거하고 제목 텍스트만 남긴다(편집자가 InDesign에서 자유롭게 재구성).
    """

    out: List[Dict[str, Any]] = []
    for blk in blocks:
        if blk.get("t") != "Header":
            out.append(blk)
            continue

        level, attr, inlines = blk.get("c", [None, None, None])
        text = _header_text(blk) or ""
        clean = _DATABOOK_HEADING_NUM_PREFIX_RE.sub("", text).strip()
        if clean != text and clean:
            # '2.4.1 1. ...' → '...'
            clean = _DATABOOK_HEADING_STEP_PREFIX_RE.sub("", clean).strip() or clean
            out.append({"t": "Header", "c": [level, attr, _str_para(clean).get("c", [])]})
            continue

        out.append(blk)

    return out


def _drop_internal_references_sections(blocks: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    i = 0
    while i < len(blocks):
        blk = blocks[i]
        if blk.get("t") == "Header":
            text = _header_text(blk) or ""
            if "참고문헌" in text or "References" in text:
                i = _section_end(blocks, i)
                continue
        out.append(blk)
        i += 1
    return out


def _para_has_image(block: Dict[str, Any]) -> bool:
    if block.get("t") != "Para":
        return False
    for el in block.get("c", []):
        if el.get("t") == "Image":
            return True
    return False


def _para_text(block: Dict[str, Any]) -> str:
    if block.get("t") != "Para":
        return ""
    return _normalize_text(_inlines_to_text(block.get("c", [])))


def _is_figure_pair(blocks: Sequence[Dict[str, Any]], i: int) -> bool:
    if i + 1 >= len(blocks):
        return False
    return blocks[i].get("t") == "Para" and _para_has_image(blocks[i]) and blocks[i + 1].get("t") == "Para" and not _para_has_image(blocks[i + 1])


def _split_caption_base(caption: str) -> Tuple[str, str]:
    base, sep, tail = caption.partition(":")
    if not sep:
        return caption.strip(), ""
    return base.strip(), tail.strip()


@dataclass
class MovedFigure:
    section: str
    base: str
    image_para: Dict[str, Any]
    caption_para: Dict[str, Any]


def _collapse_carousel_figures(blocks: List[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], List[MovedFigure]]:
    moved: List[MovedFigure] = []
    out: List[Dict[str, Any]] = []

    current_section = ""
    i = 0
    while i < len(blocks):
        blk = blocks[i]
        if blk.get("t") == "Header":
            current_section = _header_text(blk) or current_section
            out.append(blk)
            i += 1
            continue

        if not _is_figure_pair(blocks, i):
            out.append(blk)
            i += 1
            continue

        caption = _para_text(blocks[i + 1])
        base, variant = _split_caption_base(caption)

        # Look ahead for consecutive figure-pairs with the same base caption.
        pairs: List[Tuple[Dict[str, Any], Dict[str, Any], str, str]] = []
        j = i
        while j < len(blocks) and _is_figure_pair(blocks, j):
            cap = _para_text(blocks[j + 1])
            b, v = _split_caption_base(cap)
            if b != base:
                break
            pairs.append((blocks[j], blocks[j + 1], b, v))
            j += 2

        if len(pairs) <= 1:
            out.append(blocks[i])
            out.append(blocks[i + 1])
            i += 2
            continue

        keep_idx = 0
        for k, (_img, _cap, _b, v) in enumerate(pairs):
            if "전영역" in v.replace(" ", ""):
                keep_idx = k
                break

        keep_img, keep_cap, _b, _v = pairs[keep_idx]
        out.append(keep_img)
        out.append(keep_cap)
        out.append(_str_para("※ 동일 지표의 세부 도표(영역별/기관별)는 부록으로 이동"))

        for k, (img, cap, b, _v2) in enumerate(pairs):
            if k == keep_idx:
                continue
            moved.append(MovedFigure(section=current_section, base=b, image_para=img, caption_para=cap))

        i = j

    return out, moved


def _make_figure(image_para: Dict[str, Any], caption_text: str) -> Dict[str, Any]:
    caption_inlines = _str_para(caption_text).get("c", [])
    image_inlines = image_para.get("c", [])
    return {
        "t": "Figure",
        "c": [
            ["", [], []],
            [None, [{"t": "Plain", "c": caption_inlines}]],
            [{"t": "Plain", "c": image_inlines}],
        ],
    }


def _set_table_caption(table: Dict[str, Any], caption_text: str) -> Dict[str, Any]:
    if table.get("t") != "Table":
        return table
    c = table.get("c")
    if not isinstance(c, list) or len(c) < 2:
        return table
    caption_inlines = _str_para(caption_text).get("c", [])
    c[1] = [None, [{"t": "Plain", "c": caption_inlines}]]
    table["c"] = c
    return table


_FIG_PREFIX_RE = re.compile(r"^(?:그림|Figure|Fig\.)\s*[A-Z0-9]+(?:[.-][0-9]+)*\s*[:.)-]?\s*", re.IGNORECASE)
_TAB_PREFIX_RE = re.compile(r"^(?:표|Table)\s*[A-Z0-9]+(?:[.-][0-9]+)*\s*[:.)-]?\s*", re.IGNORECASE)


def _strip_caption_prefix(text: str, *, kind: str) -> str:
    text = _normalize_text(text)
    if kind == "figure":
        text = _FIG_PREFIX_RE.sub("", text)
    else:
        text = _TAB_PREFIX_RE.sub("", text)
    return text.strip().lstrip("-").strip()


_REF_RE = re.compile(r"(그림|표)\s*([A-Z0-9]+(?:[.-][0-9]+)*)")


def _replace_refs_in_inlines(
    inlines: List[Dict[str, Any]],
    *,
    fig_map: Dict[str, str],
    tab_map: Dict[str, str],
) -> List[Dict[str, Any]]:
    def replace_in_str(s: str) -> str:
        s = s.replace("\u00a0", " ")

        def repl(m: re.Match[str]) -> str:
            prefix = m.group(1)
            old_id = m.group(2)
            if prefix == "그림" and old_id in fig_map:
                return f"{prefix} {fig_map[old_id]}"
            if prefix == "표" and old_id in tab_map:
                return f"{prefix} {tab_map[old_id]}"
            return m.group(0)

        s = _REF_RE.sub(repl, s)
        # 표현 방식 통일(예: '표 3.2' → '표 3-2')
        s = re.sub(r"(그림|표)\s*([0-9]+)\.([0-9]+)", r"\1 \2-\3", s)
        return s

    def process_inline(el: Dict[str, Any]) -> Dict[str, Any]:
        t = el.get("t")
        if t == "Str":
            return {"t": "Str", "c": replace_in_str(el.get("c", ""))}
        if t in {"Emph", "Strong", "SmallCaps", "Strikeout", "Superscript", "Subscript"}:
            return {"t": t, "c": _replace_refs_in_inlines(el.get("c", []), fig_map=fig_map, tab_map=tab_map)}
        if t == "Span":
            attr, span_inlines = el.get("c", [None, []])
            return {
                "t": "Span",
                "c": [attr, _replace_refs_in_inlines(span_inlines or [], fig_map=fig_map, tab_map=tab_map)],
            }
        if t == "Link":
            attr, link_inlines, target = el.get("c", [None, [], None])
            return {
                "t": "Link",
                "c": [attr, _replace_refs_in_inlines(link_inlines or [], fig_map=fig_map, tab_map=tab_map), target],
            }
        if t == "Image":
            attr, alt_inlines, target = el.get("c", [None, [], None])
            return {
                "t": "Image",
                "c": [attr, _replace_refs_in_inlines(alt_inlines or [], fig_map=fig_map, tab_map=tab_map), target],
            }
        return el

    out: List[Dict[str, Any]] = []
    i = 0
    while i < len(inlines):
        el = process_inline(inlines[i])
        out.append(el)

        # Handle split tokens: Str("그림") Space Str("2.1는")
        if (
            i + 2 < len(inlines)
            and inlines[i].get("t") == "Str"
            and inlines[i].get("c") in {"그림", "표"}
            and inlines[i + 1].get("t") == "Space"
            and inlines[i + 2].get("t") == "Str"
        ):
            prefix = inlines[i].get("c")
            token = inlines[i + 2].get("c", "").replace("\u00a0", "")
            m = re.match(r"^([A-Z0-9]+(?:[.-][0-9]+)*)(.*)$", token)
            if m:
                old_id, suffix = m.group(1), m.group(2)
                new_id = fig_map.get(old_id) if prefix == "그림" else tab_map.get(old_id)
                if new_id:
                    out[-1] = {"t": "Str", "c": prefix}
                    out.append({"t": "Space"})
                    out.append({"t": "Str", "c": f"{new_id}{suffix}"})
                    i += 3
                    continue
                # 매핑이 없더라도 표현 방식(점→하이픈)은 통일한다.
                if "." in old_id and "-" not in old_id:
                    normalized = old_id.replace(".", "-", 1)
                    out[-1] = {"t": "Str", "c": prefix}
                    out.append({"t": "Space"})
                    out.append({"t": "Str", "c": f"{normalized}{suffix}"})
                    i += 3
                    continue
        i += 1

    return out


def _walk_blocks(blocks: List[Dict[str, Any]]) -> Iterator[Dict[str, Any]]:
    for blk in blocks:
        yield blk
        if blk.get("t") == "Table":
            # Walk into table cells to apply URL/ref cleanup if needed.
            c = blk.get("c", [])
            if not isinstance(c, list) or len(c) < 5:
                continue
            # TableHead, TableBodies, TableFoot contain nested blocks
            for container in c[3:]:
                yield from _walk_table_container(container)
        elif blk.get("t") in {"BulletList", "OrderedList"}:
            items = blk.get("c", [])
            if blk.get("t") == "OrderedList":
                items = items[1]
            for item in items:
                yield from _walk_blocks(item)
        elif blk.get("t") == "Figure":
            # Figure contents is list of blocks at c[2]
            c = blk.get("c", [])
            if isinstance(c, list) and len(c) >= 3 and isinstance(c[2], list):
                yield from _walk_blocks(c[2])


def _walk_table_container(container: Any) -> Iterator[Dict[str, Any]]:
    if not isinstance(container, list):
        return
    # TableHead: [Attr, [Row]]
    # TableBody: [Attr, RowHeadCols, [Row], [Row]]
    # TableFoot: [Attr, [Row]]
    for el in container:
        if isinstance(el, list):
            yield from _walk_table_container(el)
        elif isinstance(el, dict) and el.get("t") == "Row":
            # Row: [Attr, [Cell]]
            for cell in el.get("c", [None, []])[1]:
                yield from _walk_table_container(cell)
        elif isinstance(el, dict) and el.get("t") == "Cell":
            # Cell: [Attr, Alignment, RowSpan, ColSpan, [Block]]
            yield from _walk_blocks(el.get("c", [None, None, None, None, []])[4])


_URL_RE = re.compile(r"https?://\\S+")


def _replace_urls_in_inlines(inlines: List[Dict[str, Any]], replacement: str) -> List[Dict[str, Any]]:
    def replace_str(s: str) -> str:
        return _URL_RE.sub(replacement, s)

    out: List[Dict[str, Any]] = []
    for el in inlines:
        t = el.get("t")
        if t == "Str":
            out.append({"t": "Str", "c": replace_str(el.get("c", ""))})
            continue
        if t == "Link":
            attr, link_inlines, target = el.get("c", [None, None, None])
            out.append({"t": "Link", "c": [attr, _replace_urls_in_inlines(link_inlines or [], replacement), target]})
            continue
        if t in {"Emph", "Strong"}:
            out.append({"t": t, "c": _replace_urls_in_inlines(el.get("c", []), replacement)})
            continue
        if t == "Span":
            attr, span_inlines = el.get("c", [None, []])
            out.append({"t": "Span", "c": [attr, _replace_urls_in_inlines(span_inlines or [], replacement)]})
            continue
        out.append(el)
    return out


@dataclass
class NumberingContext:
    label: str
    fig: int = 0
    tab: int = 0
    fig_map: Dict[str, str] = field(default_factory=dict)
    tab_map: Dict[str, str] = field(default_factory=dict)


def _number_and_restructure(blocks: List[Dict[str, Any]], *, ctx: NumberingContext) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    i = 0
    while i < len(blocks):
        if _is_figure_pair(blocks, i):
            img = blocks[i]
            cap = blocks[i + 1]
            original = _para_text(cap)

            # Extract existing id if present (e.g., 그림 2.1:)
            m = re.search(r"(?:^|\b)(?:그림|Figure|Fig\.)\s*([A-Z0-9]+(?:[.-][0-9]+)*)", original, re.IGNORECASE)
            old_id = m.group(1) if m else ""

            ctx.fig += 1
            new_id = f"{ctx.label}-{ctx.fig}"
            if old_id:
                ctx.fig_map[old_id] = new_id

            clean = _strip_caption_prefix(original, kind="figure")
            caption = f"그림 {new_id}. {clean}" if clean else f"그림 {new_id}."
            out.append(_make_figure(img, caption))
            i += 2
            continue

        if blocks[i].get("t") == "Para" and i + 1 < len(blocks) and blocks[i + 1].get("t") == "Table":
            cap_text = _para_text(blocks[i])
            if re.match(r"^(?:표\s*[0-9A-Z]|Table\b|Table:)", cap_text) and len(cap_text) <= 180:
                m = re.search(r"(?:^|\b)(?:표|Table)\s*([A-Z0-9]+(?:[.-][0-9]+)*)", cap_text, re.IGNORECASE)
                old_id = m.group(1) if m else ""

                ctx.tab += 1
                new_id = f"{ctx.label}-{ctx.tab}"
                if old_id:
                    ctx.tab_map[old_id] = new_id

                clean = _strip_caption_prefix(cap_text, kind="table")
                caption = f"표 {new_id}. {clean}" if clean else f"표 {new_id}."
                table = json.loads(json.dumps(blocks[i + 1]))
                out.append(_set_table_caption(table, caption))
                i += 2
                continue

        out.append(blocks[i])
        i += 1

    # Second pass: update any internal references (그림 2.1 → 그림 3-1)
    for blk in _walk_blocks(out):
        if blk.get("t") == "Para":
            blk["c"] = _replace_refs_in_inlines(blk.get("c", []), fig_map=ctx.fig_map, tab_map=ctx.tab_map)
    return out


def _append_moved_figures_as_appendix(
    blocks: List[Dict[str, Any]], *, moved: List[MovedFigure], header_level: int
) -> List[Dict[str, Any]]:
    if not moved:
        return blocks

    out = list(blocks)
    out.append(
        {
            "t": "Header",
            "c": [header_level, ["appendix-supplement-figures", [], []], _str_para("추가 그림(캐러셀 해제본)").get("c", [])],
        }
    )

    def key(m: MovedFigure) -> Tuple[str, str]:
        return (m.section, m.base)

    moved_sorted = sorted(moved, key=key)
    current_section = None
    current_base = None
    for item in moved_sorted:
        if item.section and item.section != current_section:
            current_section = item.section
            current_base = None
            out.append({"t": "Header", "c": [header_level + 1, ["", [], []], _str_para(current_section).get("c", [])]})
        if item.base and item.base != current_base:
            current_base = item.base
            out.append({"t": "Header", "c": [header_level + 2, ["", [], []], _str_para(current_base).get("c", [])]})
        out.append(item.image_para)
        out.append(item.caption_para)

    return out


def _cell(text: str) -> List[Any]:
    return [
        ["", [], []],
        {"t": "AlignDefault"},
        1,
        1,
        [{"t": "Plain", "c": _str_para(text).get("c", [])}],
    ]


def _row(cells: Sequence[str]) -> List[Any]:
    return [["", [], []], [_cell(c) for c in cells]]


def _make_table(rows: List[List[str]], *, col_widths: Optional[List[float]] = None) -> Dict[str, Any]:
    if not rows:
        raise ValueError("rows must not be empty")
    n_cols = len(rows[0])
    if any(len(r) != n_cols for r in rows):
        raise ValueError("all rows must have the same number of columns")

    if col_widths is None:
        col_widths = [1.0 / n_cols] * n_cols
    total = sum(col_widths)
    if total <= 0:
        col_widths = [1.0 / n_cols] * n_cols
        total = 1.0
    col_widths = [w / total for w in col_widths]

    colspecs = [[{"t": "AlignDefault"}, {"t": "ColWidth", "c": w}] for w in col_widths]
    pandoc_rows = [_row(r) for r in rows]

    # TableHead/TableFoot는 비워두고, 모든 행을 bodyRows에 넣는다(Word 변환 안정성 우선).
    return {
        "t": "Table",
        "c": [
            ["", [], []],
            [None, []],
            colspecs,
            [["", [], []], []],
            [[["", [], []], 0, [], pandoc_rows]],
            [["", [], []], []],
        ],
    }


def _try_parse_rank_institute_table(lines: List[str]) -> Optional[List[List[str]]]:
    """
    Word 변환 과정에서 '순위/기관/N_docs' 표가 줄 단위 문단으로 깨진 경우를
    3열 테이블로 복구한다.
    """

    if len(lines) < 6:
        return None
    if lines[:3] != ["순위", "기관", "N_docs"]:
        return None

    data = lines[3:]
    if not data:
        return None

    rows: List[List[str]] = []
    if len(data) % 3 == 0:
        for i in range(0, len(data), 3):
            rows.append(data[i : i + 3])
    else:
        # Fallback: rank(숫자) → 기관(문자열) → N_docs(숫자) 형태로 스캔.
        i = 0
        num_re = re.compile(r"^\d+(?:,\d{3})*$")
        while i < len(data):
            rank = data[i].strip()
            if not num_re.match(rank):
                return None
            i += 1

            inst_parts: List[str] = []
            while i < len(data) and not num_re.match(data[i].strip()):
                inst_parts.append(data[i].strip())
                i += 1
            if i >= len(data):
                return None

            n_docs = data[i].strip()
            i += 1
            rows.append([rank, " ".join(inst_parts).strip(), n_docs])

    return [lines[:3]] + rows


def _fix_appendix_top_institutes_tables(blocks: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    상세 리포트(Appendix C/E/G 등)에서 '4-2. 상위 기관' 표가 DOCX 출력에서
    표가 아니라 문단으로 깨지는 케이스가 있어, 패턴 매칭으로 Word Table로 복구한다.
    """

    out: List[Dict[str, Any]] = []
    i = 0
    while i < len(blocks):
        blk = blocks[i]
        if blk.get("t") != "Header":
            out.append(blk)
            i += 1
            continue

        text = _header_text(blk) or ""
        if "4-2." not in text or "상위 기관" not in text or "문헌 수 기준" not in text:
            out.append(blk)
            i += 1
            continue

        # Process the whole '상위 기관' section at once.
        section_start = i
        section_end = _section_end(blocks, section_start)
        out.append(blk)
        section = blocks[section_start + 1 : section_end]

        markers = {"전체(Top500 + BIPM) Top5", "BIPM 한정 Top5"}

        j = 0
        while j < len(section):
            b = section[j]
            if b.get("t") == "Para":
                para_text = _para_text(b)
                if para_text in markers:
                    out.append(b)

                    # Capture subsequent Para blocks until next marker/header.
                    k = j + 1
                    captured_blocks: List[Dict[str, Any]] = []
                    captured_lines: List[str] = []
                    while k < len(section):
                        nb = section[k]
                        if nb.get("t") == "Header":
                            break
                        if nb.get("t") != "Para":
                            break
                        t = _para_text(nb)
                        if t in markers:
                            break
                        captured_blocks.append(nb)
                        if t:
                            captured_lines.append(t)
                        k += 1

                    table_rows = _try_parse_rank_institute_table(captured_lines)
                    if table_rows:
                        out.append(_make_table(table_rows, col_widths=[0.12, 0.73, 0.15]))
                        j = k
                        continue

            out.append(b)
            j += 1

        i = section_end

    return out


def _make_header(level: int, text: str, *, identifier: str = "") -> Dict[str, Any]:
    return {
        "t": "Header",
        "c": [
            level,
            [identifier, [], []],
            _str_para(text).get("c", []),
        ],
    }


def _promote_intro_paras_to_headers(blocks: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    최종보고서 원안(HWPX→DOCX)에서 '서론' 구간은 Heading 스타일이 아닌 평문으로
    들어오는 경우가 있어, 인쇄/편집용 단일원본에서 구조를 유지하도록 Header로 승격한다.

    - '서 론' → Heading 1
    - 첫 '연구개발의 필요성' → Heading 2
    - '연구개발의 목적'/'연구개발의 필요성'(2번째)/'연구개발의 범위' → Heading 3
    """

    first_header_idx = next((i for i, b in enumerate(blocks) if b.get("t") == "Header"), None)
    if first_header_idx is None:
        return blocks

    intro_idx = None
    for i in range(first_header_idx):
        if blocks[i].get("t") == "Para" and _para_text(blocks[i]) == "서 론":
            intro_idx = i
            break
    if intro_idx is None:
        return blocks

    need_idxs: List[int] = []
    purpose_idx: Optional[int] = None
    scope_idx: Optional[int] = None
    for i in range(intro_idx + 1, first_header_idx):
        if blocks[i].get("t") != "Para":
            continue
        t = _para_text(blocks[i])
        if t == "연구개발의 필요성":
            need_idxs.append(i)
        elif t == "연구개발의 목적" and purpose_idx is None:
            purpose_idx = i
        elif t == "연구개발의 범위" and scope_idx is None:
            scope_idx = i

    out: List[Dict[str, Any]] = []
    for i, b in enumerate(blocks):
        if i == intro_idx:
            out.append(_make_header(1, "서 론", identifier="chapter-1-intro"))
            continue
        if need_idxs and i == need_idxs[0]:
            out.append(_make_header(2, "연구개발의 필요성"))
            continue
        if purpose_idx is not None and i == purpose_idx:
            out.append(_make_header(3, "연구개발의 목적"))
            continue
        if len(need_idxs) >= 2 and i == need_idxs[1]:
            out.append(_make_header(3, "연구개발의 필요성"))
            continue
        if scope_idx is not None and i == scope_idx:
            out.append(_make_header(3, "연구개발의 범위"))
            continue
        out.append(b)

    return out


_KOREAN_ITEM_LETTERS = ["가", "나", "다", "라", "마", "바", "사", "아", "자", "차", "카", "타", "파", "하"]
_CHAPTER_PREFIX_RE = re.compile(r"^제\s*\d+\s*장\s+")
_SECTION_PREFIX_RE = re.compile(r"^제\s*\d+\s*절\s+")
_SUBSECTION_PREFIX_RE = re.compile(r"^\d+\.\s+")
_KOREAN_ITEM_PREFIX_RE = re.compile(r"^[가나다라마바사아자차카타파하](?:\d+)?\.\s+")
_KOREAN_SUBITEM_PREFIX_RE = re.compile(r"^\d+\)\s+")
_BRACKET_FIG_RE = re.compile(r"^\[\s*그림\s*\]\s*")


def _korean_item_letter(idx: int) -> str:
    if idx <= 0:
        raise ValueError("idx must be >= 1")
    base_len = len(_KOREAN_ITEM_LETTERS)
    if idx <= base_len:
        return _KOREAN_ITEM_LETTERS[idx - 1]
    q, r = divmod(idx - 1, base_len)
    return f"{_KOREAN_ITEM_LETTERS[r]}{q + 1}"


def _strip_heading_prefix(text: str, *, level: int) -> str:
    text = _normalize_text(text)
    if level == 1:
        text = _CHAPTER_PREFIX_RE.sub("", text)
    elif level == 2:
        text = _SECTION_PREFIX_RE.sub("", text)
    elif level == 3:
        text = _SUBSECTION_PREFIX_RE.sub("", text)
    elif level == 4:
        text = _KOREAN_ITEM_PREFIX_RE.sub("", text)
    elif level == 5:
        text = _KOREAN_SUBITEM_PREFIX_RE.sub("", text)
    return text.strip()


def _apply_korean_report_heading_numbering(blocks: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    한글(HWP) 최종보고서 관행에 맞춰 Heading 텍스트에 번호를 붙인다.
    - Heading 1: '제 n 장 ...'
    - Heading 2: '제 n 절 ...'
    - Heading 3: 'n. ...'
    - Heading 4: '가. ...'
    - Heading 5: 'n) ...'

    부록(Heading 1 '부록') 이후는 Appendix 자체 번호체계를 유지한다.
    """

    out: List[Dict[str, Any]] = []
    in_main = True
    chapter = 0
    section = 0
    subsection = 0
    item = 0
    subitem = 0

    for blk in blocks:
        if blk.get("t") != "Header":
            out.append(blk)
            continue

        level, attr, _inlines = blk.get("c", [None, None, None])
        if not isinstance(level, int):
            out.append(blk)
            continue

        text = _header_text(blk) or ""
        if level == 1 and _normalize_text(text) == "부록":
            in_main = False
            out.append(blk)
            continue

        if not in_main:
            out.append(blk)
            continue

        title = _strip_heading_prefix(text, level=level)

        if level == 1:
            chapter += 1
            section = 0
            subsection = 0
            item = 0
            subitem = 0
            new_text = f"제 {chapter} 장 {title}"
        elif level == 2:
            section += 1
            subsection = 0
            item = 0
            subitem = 0
            new_text = f"제 {section} 절 {title}"
        elif level == 3:
            subsection += 1
            item = 0
            subitem = 0
            new_text = f"{subsection}. {title}"
        elif level == 4:
            item += 1
            subitem = 0
            new_text = f"{_korean_item_letter(item)}. {title}"
        elif level == 5:
            subitem += 1
            new_text = f"{subitem}) {title}"
        else:
            out.append(blk)
            continue

        out.append({"t": "Header", "c": [level, attr, _str_para(new_text).get("c", [])]})

    return out


def _number_bracket_figure_tables(blocks: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    최종보고서 원안에는 본문에 '[그림 ] ...' 형태의 placeholder 캡션이
    표(table) 내부(2x1 레이아웃)로 들어간 경우가 있다. PDF/인쇄본에서
    '그림번호가 빠진 것'처럼 보이므로 장-번호 형식으로 채운다.

    예: '[그림 ] 미-중 ...' → '그림 1-1. 미-중 ...'
    """

    def update_first_cell_caption(table: Dict[str, Any], *, fig_id: str) -> Tuple[Dict[str, Any], bool]:
        if table.get("t") != "Table":
            return table, False
        c = table.get("c", [])
        if not isinstance(c, list) or len(c) < 5:
            return table, False
        bodies = c[4]
        if not isinstance(bodies, list) or not bodies:
            return table, False
        body0 = bodies[0]
        if not isinstance(body0, list) or len(body0) < 4:
            return table, False
        rows = body0[3]
        if not isinstance(rows, list) or not rows:
            return table, False
        first_row = rows[0]
        if not isinstance(first_row, list) or len(first_row) < 2:
            return table, False
        cells = first_row[1]
        if not isinstance(cells, list) or not cells:
            return table, False
        first_cell = cells[0]
        if not isinstance(first_cell, list) or len(first_cell) < 5:
            return table, False
        cell_blocks = first_cell[4]
        if not isinstance(cell_blocks, list) or not cell_blocks:
            return table, False

        for bi, blk in enumerate(cell_blocks):
            if not isinstance(blk, dict) or blk.get("t") not in {"Para", "Plain"}:
                continue
            text = _normalize_text(_inlines_to_text(blk.get("c", [])))
            if not text.startswith("[그림"):
                continue
            clean = _BRACKET_FIG_RE.sub("", text).strip()
            if not clean:
                clean = text
            new_caption = f"그림 {fig_id}. {clean}"
            cell_blocks[bi] = {"t": blk.get("t"), "c": _str_para(new_caption).get("c", [])}
            first_cell[4] = cell_blocks
            cells[0] = first_cell
            first_row[1] = cells
            rows[0] = first_row
            body0[3] = rows
            bodies[0] = body0
            c[4] = bodies
            table["c"] = c
            return table, True

        return table, False

    out: List[Dict[str, Any]] = []
    in_main = True
    chapter = 0
    fig_counter = 0

    for blk in blocks:
        if blk.get("t") == "Header":
            level = blk.get("c", [None])[0]
            text = _header_text(blk) or ""
            if level == 1:
                if text == "부록":
                    in_main = False
                if in_main:
                    chapter += 1
                    fig_counter = 0
            out.append(blk)
            continue

        if in_main and blk.get("t") == "Table" and chapter > 0:
            fig_counter += 1
            fig_id = f"{chapter}-{fig_counter}"
            table = json.loads(json.dumps(blk))
            table, changed = update_first_cell_caption(table, fig_id=fig_id)
            if changed:
                out.append(table)
                continue
            fig_counter -= 1

        out.append(blk)

    return out


def merge_docx_master(
    *,
    master_docx: Path,
    databook_docx: Path,
    output_docx: Path,
    url_replacement: str,
) -> None:
    pandoc = _ensure_pandoc()
    master_docx = master_docx.resolve()
    databook_docx = databook_docx.resolve()
    output_docx = output_docx.resolve()
    output_docx.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="kriss_docx_merge_") as tmp:
        tmp = Path(tmp)
        master_media = tmp / "master_media"
        databook_media = tmp / "databook_media"
        master_json = tmp / "master.json"
        databook_json = tmp / "databook.json"

        _run([pandoc, str(master_docx), "--extract-media", str(master_media), "-t", "json", "-o", str(master_json)])
        _run([pandoc, str(databook_docx), "--extract-media", str(databook_media), "-t", "json", "-o", str(databook_json)])

        master = _load_json(master_json)
        databook = _load_json(databook_json)

        # Pandoc은 meta.title이 있으면 DOCX 상단에 Title 문단을 생성한다.
        # HWPX 변환 파이프라인에서 들어온 "Converted HWPX" 타이틀은 편집에 불필요하므로 제거한다.
        if isinstance(master.get("meta"), dict):
            master["meta"].pop("title", None)

        master_blocks: List[Dict[str, Any]] = master.get("blocks", [])
        databook_blocks: List[Dict[str, Any]] = databook.get("blocks", [])

        # 1) 데이터북(최신)으로 덮어쓸 구간: 최종보고서 3장 내부 3개 절
        mapping = [
            (
                "표준과학 연구 영역 분석을 위한 문헌 집합",
                "2. 문헌 집합 및 연구지형도 구축",
            ),
            (
                "기관 연구역량 수준 진단을 위한 과학계량지표 빅데이터 분석",
                "3. 기관 연구역량 진단",
            ),
            (
                "미래 연구영역 전략수립 지원을 위한 유망연구영역 도출",
                "4. 유망 연구영역 도출",
            ),
        ]

        for master_header, databook_chapter in mapping:
            master_header_idx = _find_header_index(master_blocks, needle=master_header)
            content_start = master_header_idx + 1
            content_end = _section_end(master_blocks, master_header_idx)

            chapter = _extract_section_by_header(databook_blocks, header_text=databook_chapter, level=1)
            # Drop chapter's own Heading 1 and indent sub-structure by 1 level.
            chapter = chapter[1:]
            chapter = _drop_internal_references_sections(chapter)
            chapter = _shift_heading_levels(chapter, +1)
            chapter = _strip_databook_heading_numbers(chapter)

            master_blocks = master_blocks[:content_start] + chapter + master_blocks[content_end:]

        # 2) 캐러셀(유사 그림 묶음) 축약: 동일 base caption 반복은 하나만 남기고 부록으로 이동
        chapter3_idx = _find_header_index(master_blocks, needle="연구개발수행 내용 및 결과", level=1)
        chapter3_end = _section_end(master_blocks, chapter3_idx)
        chapter3 = [json.loads(json.dumps(b)) for b in master_blocks[chapter3_idx:chapter3_end]]
        chapter3_collapsed, moved = _collapse_carousel_figures(chapter3)
        master_blocks = master_blocks[:chapter3_idx] + chapter3_collapsed + master_blocks[chapter3_end:]

        # 3) 부록 추가: (a) 캐러셀에서 이동한 그림 (b) 데이터북 부록 전체
        appendix_start = _find_header_index(databook_blocks, needle="Appendix A", level=1, contains=True)
        appendices = [json.loads(json.dumps(b)) for b in databook_blocks[appendix_start:]]
        appendices = _shift_heading_levels(appendices, +1)
        appendices = _fix_appendix_top_institutes_tables(appendices)

        master_blocks.append({"t": "Header", "c": [1, ["appendix", [], []], _str_para("부록").get("c", [])]})
        master_blocks = _append_moved_figures_as_appendix(master_blocks, moved=moved, header_level=2)
        master_blocks.extend(appendices)

        # 4) URL 정책: 인쇄본 본문에서는 URL을 '데이터북 참조'로 통일 (최종 참고문헌 구간만 유지)
        refs_start = None
        refs_end = None
        try:
            refs_start = _find_header_index(master_blocks, needle="참고문헌", level=1)
            refs_end = _section_end(master_blocks, refs_start)
        except ValueError:
            refs_start = None
            refs_end = None

        for idx, blk in enumerate(master_blocks):
            if refs_start is not None and refs_end is not None and refs_start <= idx < refs_end:
                continue
            if blk.get("t") == "Para":
                blk["c"] = _replace_urls_in_inlines(blk.get("c", []), url_replacement)

        # 5) 그림/표 번호 재부여: 3장(라벨 '3') + 부록(라벨 'S' 및 Appendix letter)
        chapter3_idx = _find_header_index(master_blocks, needle="연구개발수행 내용 및 결과", level=1)
        chapter3_end = _section_end(master_blocks, chapter3_idx)
        ctx_ch3 = NumberingContext(label="3")
        master_blocks[chapter3_idx:chapter3_end] = _number_and_restructure(
            master_blocks[chapter3_idx:chapter3_end], ctx=ctx_ch3
        )

        appendix_idx = _find_header_index(master_blocks, needle="부록", level=1)
        appendix_end = len(master_blocks)
        appendix_blocks = [json.loads(json.dumps(b)) for b in master_blocks[appendix_idx:appendix_end]]

        out_app: List[Dict[str, Any]] = []
        i = 0
        while i < len(appendix_blocks):
            blk = appendix_blocks[i]

            # Supplementary moved figures section: label 'S'
            if blk.get("t") == "Header" and blk.get("c", [None])[0] == 2:
                if (_header_text(blk) or "") == "추가 그림(캐러셀 해제본)":
                    end = _section_end(appendix_blocks, i)
                    ctx_supp = NumberingContext(label="S")
                    out_app.extend(_number_and_restructure(appendix_blocks[i:end], ctx=ctx_supp))
                    i = end
                    continue

                # Appendix letter sections: label 'A'/'B'/...
                text = _header_text(blk) or ""
                m = re.match(r"Appendix\s+([A-Z])\b", text)
                if m:
                    letter = m.group(1)
                    end = _section_end(appendix_blocks, i)
                    ctx = NumberingContext(label=letter)
                    out_app.extend(_number_and_restructure(appendix_blocks[i:end], ctx=ctx))
                    i = end
                    continue

            out_app.append(blk)
            i += 1

        master_blocks[appendix_idx:appendix_end] = out_app

        # 6) 한글 보고서 형식(장/절/1./가.) 복원: 서론 구간 승격 + Heading 텍스트 번호 부여
        master_blocks = _promote_intro_paras_to_headers(master_blocks)
        master_blocks = _apply_korean_report_heading_numbering(master_blocks)
        master_blocks = _number_bracket_figure_tables(master_blocks)

        master["pandoc-api-version"] = PANDOC_API_VERSION
        master["blocks"] = master_blocks

        merged_json = tmp / "merged.json"
        _write_json(merged_json, master)
        _run([pandoc, str(merged_json), "-f", "json", "-t", "docx", "-o", str(output_docx)])


def main() -> int:
    parser = argparse.ArgumentParser(description="KRISS 최종보고서(DOCX) + 데이터북(DOCX) 머지(인쇄/편집용).")
    parser.add_argument("--master-docx", type=Path, required=True, help="최종보고서 DOCX(단일 원본) 경로")
    parser.add_argument("--databook-docx", type=Path, required=True, help="데이터북 DOCX(최신) 경로")
    parser.add_argument("--output", type=Path, required=True, help="머지 결과 DOCX 경로")
    parser.add_argument(
        "--url-replacement",
        type=str,
        default="(데이터북 참조)",
        help="본문에서 URL을 대체할 문구(기본: (데이터북 참조))",
    )
    args = parser.parse_args()

    merge_docx_master(
        master_docx=args.master_docx,
        databook_docx=args.databook_docx,
        output_docx=args.output,
        url_replacement=args.url_replacement,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
