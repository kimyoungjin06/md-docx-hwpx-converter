#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, Iterator, List, Optional, Sequence, Tuple

from lxml import etree


HWPX_NS = {
    "ha": "http://www.hancom.co.kr/hwpml/2011/app",
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hp10": "http://www.hancom.co.kr/hwpml/2016/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
    "hh": "http://www.hancom.co.kr/hwpml/2011/head",
    "hhs": "http://www.hancom.co.kr/hwpml/2011/history",
    "hm": "http://www.hancom.co.kr/hwpml/2011/master-page",
    "hpf": "http://www.hancom.co.kr/schema/2011/hpf",
    "dc": "http://purl.org/dc/elements/1.1/",
    "opf": "http://www.idpf.org/2007/opf/",
    "ooxmlchart": "http://www.hancom.co.kr/hwpml/2016/ooxmlchart",
    "hwpunitchar": "http://www.hancom.co.kr/hwpml/2016/HwpUnitChar",
    "epub": "http://www.idpf.org/2007/ops",
    "config": "urn:oasis:names:tc:opendocument:xmlns:config:1.0",
}


DEFAULT_TEMPLATE_HWPX = (
    Path(__file__).resolve().parents[1]
    / "빅데이터_기반_표준연구_기관_성과분석_및_유망연구영역_도출_최종보고서.hwpx"
)


def _run(cmd: Sequence[str], *, cwd: Optional[Path] = None) -> None:
    subprocess.run(list(cmd), check=True, cwd=str(cwd) if cwd else None)


def _ensure_pandoc() -> str:
    pandoc = shutil.which("pandoc")
    if not pandoc:
        raise RuntimeError("pandoc is required but was not found in PATH.")
    return pandoc


def _local_name(tag: str) -> str:
    if "}" in tag:
        return tag.rsplit("}", 1)[-1]
    return tag


@dataclass(frozen=True)
class HwpStyle:
    style_id: str
    para_pr_id: str
    char_pr_id: str
    name: str
    eng_name: str


def _load_hwpx_styles(header_xml: bytes) -> Dict[str, HwpStyle]:
    root = etree.fromstring(header_xml)  # noqa: S320 (trusted local file)
    styles: Dict[str, HwpStyle] = {}
    for el in root.xpath('//*[local-name()="style"]'):
        style_id = el.get("id")
        if not style_id:
            continue
        styles[style_id] = HwpStyle(
            style_id=style_id,
            para_pr_id=el.get("paraPrIDRef") or "0",
            char_pr_id=el.get("charPrIDRef") or "0",
            name=el.get("name") or "",
            eng_name=el.get("engName") or "",
        )
    return styles


def _normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text.replace("\u00a0", " ")).strip()


def _inlines_to_text(inlines: Sequence[Dict[str, Any]]) -> str:
    parts: List[str] = []
    for el in inlines:
        t = el.get("t")
        if t == "Str":
            parts.append(el.get("c", ""))
        elif t in {"Space", "SoftBreak"}:
            parts.append(" ")
        elif t == "LineBreak":
            parts.append("\n")
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
            c = el.get("c", [None, [], ["", ""]])
            alt = _normalize_text(_inlines_to_text(c[1] if isinstance(c, list) and len(c) >= 2 else []))
            target = c[2] if isinstance(c, list) and len(c) >= 3 else ["", ""]
            url = target[0] if isinstance(target, list) and len(target) >= 1 else ""
            name = Path(url).name if url else ""
            if name and alt:
                parts.append(f"[이미지: {name} {alt}]")
            elif name:
                parts.append(f"[이미지: {name}]")
            elif alt:
                parts.append(f"[이미지: {alt}]")
            else:
                parts.append("[이미지]")
        else:
            parts.append(" ")
    return "".join(parts)


@dataclass(frozen=True)
class DocPara:
    kind: str  # "p" | "h"
    text: str
    level: int = 0


def _blocks_to_text(blocks: Sequence[Dict[str, Any]]) -> str:
    parts: List[str] = []
    for para in _flatten_blocks(blocks):
        if para.text:
            parts.append(para.text)
    return _normalize_text(" ".join(parts))


def _table_rows(table_block: Dict[str, Any]) -> List[List[str]]:
    c = table_block.get("c")
    if not isinstance(c, list) or len(c) < 6:
        return []
    table_head = c[3]
    table_bodies = c[4]
    table_foot = c[5]

    rows: List[List[str]] = []

    def rows_from_row_objs(row_objs: Any) -> None:
        if not isinstance(row_objs, list):
            return
        for row in row_objs:
            if not isinstance(row, dict) or row.get("t") != "Row":
                continue
            row_c = row.get("c", [None, []])
            cells = row_c[1] if isinstance(row_c, list) and len(row_c) >= 2 else []
            row_texts: List[str] = []
            if not isinstance(cells, list):
                continue
            for cell in cells:
                if not isinstance(cell, dict) or cell.get("t") != "Cell":
                    row_texts.append("")
                    continue
                cell_c = cell.get("c", [None, None, None, None, []])
                blocks = cell_c[4] if isinstance(cell_c, list) and len(cell_c) >= 5 else []
                row_texts.append(_blocks_to_text(blocks if isinstance(blocks, list) else []))
            rows.append(row_texts)

    if isinstance(table_head, dict) and table_head.get("t") == "TableHead":
        head_rows = table_head.get("c", [None, []])[1]
        rows_from_row_objs(head_rows)

    if isinstance(table_bodies, list):
        for body in table_bodies:
            if not isinstance(body, dict) or body.get("t") != "TableBody":
                continue
            body_c = body.get("c", [None, None, [], []])
            head_rows = body_c[2] if isinstance(body_c, list) and len(body_c) >= 3 else []
            body_rows = body_c[3] if isinstance(body_c, list) and len(body_c) >= 4 else []
            rows_from_row_objs(head_rows)
            rows_from_row_objs(body_rows)

    if isinstance(table_foot, dict) and table_foot.get("t") == "TableFoot":
        foot_rows = table_foot.get("c", [None, []])[1]
        rows_from_row_objs(foot_rows)

    return rows


def _flatten_blocks(blocks: Sequence[Dict[str, Any]]) -> Iterator[DocPara]:
    for block in blocks:
        if not isinstance(block, dict):
            continue
        t = block.get("t")
        if t == "Header":
            level, _attr, inlines = block.get("c", [0, None, []])
            text = _normalize_text(_inlines_to_text(inlines if isinstance(inlines, list) else []))
            if text:
                yield DocPara(kind="h", level=int(level), text=text)
            continue

        if t in {"Para", "Plain"}:
            inlines = block.get("c", [])
            text = _normalize_text(_inlines_to_text(inlines if isinstance(inlines, list) else []))
            if text:
                yield DocPara(kind="p", text=text)
            continue

        if t == "LineBlock":
            lines = block.get("c", [])
            if isinstance(lines, list):
                rendered = "\n".join(_normalize_text(_inlines_to_text(line)) for line in lines if isinstance(line, list))
                rendered = rendered.strip()
                if rendered:
                    yield DocPara(kind="p", text=rendered)
            continue

        if t == "CodeBlock":
            c = block.get("c", [None, ""])
            text = c[1] if isinstance(c, list) and len(c) >= 2 else ""
            text = text.strip("\n")
            if text:
                yield DocPara(kind="p", text=text)
            continue

        if t == "BlockQuote":
            inner = block.get("c", [])
            if isinstance(inner, list):
                yield from _flatten_blocks(inner)
            continue

        if t == "Div":
            c = block.get("c", [None, []])
            inner = c[1] if isinstance(c, list) and len(c) >= 2 else []
            if isinstance(inner, list):
                yield from _flatten_blocks(inner)
            continue

        if t == "BulletList":
            items = block.get("c", [])
            if not isinstance(items, list):
                continue
            for item in items:
                if not isinstance(item, list):
                    continue
                text = _blocks_to_text(item)
                if text:
                    yield DocPara(kind="p", text=f"- {text}")
            continue

        if t == "OrderedList":
            c = block.get("c", [None, []])
            items = c[1] if isinstance(c, list) and len(c) >= 2 else []
            if not isinstance(items, list):
                continue
            for idx, item in enumerate(items, start=1):
                if not isinstance(item, list):
                    continue
                text = _blocks_to_text(item)
                if text:
                    yield DocPara(kind="p", text=f"{idx}. {text}")
            continue

        if t == "Table":
            caption = ""
            c = block.get("c", [])
            if isinstance(c, list) and len(c) >= 2:
                cap = c[1]
                if isinstance(cap, dict) and cap.get("t") == "Caption":
                    cap_blocks = cap.get("c", [None, []])[1]
                    if isinstance(cap_blocks, list):
                        caption = _blocks_to_text(cap_blocks)
            if caption:
                yield DocPara(kind="p", text=caption)
            rows = _table_rows(block)
            if not rows:
                yield DocPara(kind="p", text="[표]")
                continue
            for row in rows:
                if not any(cell.strip() for cell in row):
                    continue
                yield DocPara(kind="p", text=" | ".join(row))
            continue

        # Ignore blocks that are not representable as plain HWPX in this best-effort converter.


def _strip_korean_heading_prefix(level: int, text: str) -> str:
    if level == 1:
        return re.sub(r"^제\s*\d+\s*장\s*", "", text).strip()
    if level == 2:
        return re.sub(r"^제\s*\d+\s*절\s*", "", text).strip()
    if level == 3:
        return re.sub(r"^\d+\.\s*", "", text).strip()
    return text


def _style_for(para: DocPara, *, mode: str) -> str:
    if para.kind != "h":
        return "0"
    if mode == "report":
        return {1: "16", 2: "18", 3: "20", 4: "21"}.get(para.level, "0")
    # mode == "outline"
    return {1: "28", 2: "26", 3: "29"}.get(para.level, "0")


def _append_text(run: etree._Element, text: str) -> None:
    hp_ns = HWPX_NS["hp"]

    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = normalized.split("\n")
    for idx, line in enumerate(lines):
        t_el = etree.SubElement(run, f"{{{hp_ns}}}t")
        if line:
            t_el.text = line
        if idx != len(lines) - 1:
            etree.SubElement(run, f"{{{hp_ns}}}lineBreak")


def _new_paragraph(*, styles: Dict[str, HwpStyle], style_id: str, text: str) -> etree._Element:
    hp_ns = HWPX_NS["hp"]
    style = styles.get(style_id) or styles.get("0")
    if style is None:
        raise RuntimeError("Template HWPX is missing style 0 (바탕글).")

    p_el = etree.Element(
        f"{{{hp_ns}}}p",
        id="2147483648",
        paraPrIDRef=style.para_pr_id,
        styleIDRef=style.style_id,
        pageBreak="0",
        columnBreak="0",
        merged="0",
    )
    run = etree.SubElement(p_el, f"{{{hp_ns}}}run", charPrIDRef=style.char_pr_id)
    _append_text(run, text)

    # Layout segments: we keep a minimal/default segment so Hancom can reflow on open.
    seg_array = etree.SubElement(p_el, f"{{{hp_ns}}}linesegarray")
    etree.SubElement(
        seg_array,
        f"{{{hp_ns}}}lineseg",
        textpos="0",
        vertpos="0",
        vertsize="1100",
        textheight="1100",
        baseline="935",
        spacing="880",
        horzpos="0",
        horzsize="45352",
        flags="393216",
    )
    return p_el


def convert_docx_to_hwpx(
    *,
    docx_path: Path,
    template_hwpx: Path,
    output_hwpx: Path,
    mode: str,
    strip_heading_prefix: bool,
    keep_intermediate: bool,
) -> None:
    docx_path = docx_path.resolve()
    template_hwpx = template_hwpx.resolve()
    output_hwpx = output_hwpx.resolve()
    output_hwpx.parent.mkdir(parents=True, exist_ok=True)

    pandoc = _ensure_pandoc()

    workdir_ctx = tempfile.TemporaryDirectory(prefix="docx2hwpx_")
    try:
        workdir = Path(workdir_ctx.name)
        extract_media = workdir / "media"
        extract_media.mkdir(parents=True, exist_ok=True)
        json_path = workdir / "doc.json"

        _run(
            [
                pandoc,
                str(docx_path),
                "-t",
                "json",
                "--extract-media",
                str(extract_media),
                "-o",
                str(json_path),
            ],
            cwd=workdir,
        )

        doc = json.loads(json_path.read_text(encoding="utf-8"))
        blocks = doc.get("blocks", [])
        if not isinstance(blocks, list):
            raise RuntimeError("Unexpected pandoc JSON: missing 'blocks'.")

        with zipfile.ZipFile(template_hwpx) as zf:
            styles = _load_hwpx_styles(zf.read("Contents/header.xml"))
            section_root = etree.fromstring(zf.read("Contents/section0.xml"))  # noqa: S320 (trusted local file)

        # Keep the section properties paragraph (contains <hp:secPr>).
        sec_pr_p: Optional[etree._Element] = None
        for child in section_root:
            if _local_name(child.tag) != "p":
                continue
            if child.xpath('.//*[local-name()="secPr"]'):
                sec_pr_p = child
                break
        if sec_pr_p is None:
            raise RuntimeError("Template section0.xml does not contain a <hp:secPr> paragraph.")

        for child in list(section_root):
            if child is sec_pr_p:
                continue
            section_root.remove(child)

        for para in _flatten_blocks(blocks):
            text = para.text
            if para.kind == "h" and strip_heading_prefix:
                text = _strip_korean_heading_prefix(para.level, text)
            style_id = _style_for(para, mode=mode)
            section_root.append(_new_paragraph(styles=styles, style_id=style_id, text=text))

        section_xml = etree.tostring(
            section_root,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True,
        )

        with zipfile.ZipFile(template_hwpx) as zf_in, zipfile.ZipFile(output_hwpx, "w") as zf_out:
            for info in zf_in.infolist():
                data = zf_in.read(info.filename)
                if info.filename == "Contents/section0.xml":
                    data = section_xml
                zf_out.writestr(info, data)

        if keep_intermediate:
            intermediate_dir = output_hwpx.with_suffix(".docx2hwpx")
            if intermediate_dir.exists():
                shutil.rmtree(intermediate_dir)
            shutil.copytree(workdir, intermediate_dir)
    finally:
        workdir_ctx.cleanup()


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Best-effort DOCX → HWPX converter (preserves headings/outline, text; tables/images become text)."
    )
    parser.add_argument("docx", type=Path, help="Input .docx file")
    parser.add_argument("--output", type=Path, required=True, help="Output .hwpx path")
    parser.add_argument(
        "--template-hwpx",
        type=Path,
        default=DEFAULT_TEMPLATE_HWPX,
        help="Template .hwpx used for styles/settings (default: final_report_site/*최종보고서.hwpx).",
    )
    parser.add_argument(
        "--mode",
        choices=["outline", "report"],
        default="outline",
        help="Heading style mapping: 'outline' prioritizes outline tree, 'report' uses KRISS report styles (default: outline).",
    )
    parser.add_argument(
        "--no-strip-heading-prefix",
        action="store_true",
        help="Do not strip leading numbering text from headings (useful when mode=outline).",
    )
    parser.add_argument(
        "--keep-intermediate",
        action="store_true",
        help="Keep intermediate pandoc JSON + extracted media (suffix: .docx2hwpx).",
    )
    args = parser.parse_args()

    strip_heading_prefix = not args.no_strip_heading_prefix and args.mode == "report"

    convert_docx_to_hwpx(
        docx_path=args.docx,
        template_hwpx=args.template_hwpx,
        output_hwpx=args.output,
        mode=args.mode,
        strip_heading_prefix=strip_heading_prefix,
        keep_intermediate=args.keep_intermediate,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
