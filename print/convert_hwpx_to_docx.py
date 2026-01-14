#!/usr/bin/env python3
from __future__ import annotations

import argparse
import html
import io
import shutil
import subprocess
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set

from lxml import etree
from PIL import Image


@dataclass(frozen=True)
class StyleInfo:
    style_id: str
    name: str
    eng_name: str


def _local_name(tag: str) -> str:
    if "}" in tag:
        return tag.rsplit("}", 1)[-1]
    return tag


def _load_styles(header_xml: bytes) -> Dict[str, StyleInfo]:
    root = etree.fromstring(header_xml)  # noqa: S320 (trusted local file)
    styles: Dict[str, StyleInfo] = {}
    for el in root.xpath('//*[local-name()="style"]'):
        style_id = el.get("id")
        if not style_id:
            continue
        styles[style_id] = StyleInfo(
            style_id=style_id,
            name=el.get("name") or "",
            eng_name=el.get("engName") or "",
        )
    return styles


def _heading_level(style: Optional[StyleInfo]) -> Optional[int]:
    if style is None:
        return None

    eng = style.eng_name.strip().lower()
    name = style.name.strip()

    if eng.startswith("heading "):
        tail = eng.removeprefix("heading ").strip()
        try:
            level = int(tail)
        except ValueError:
            level = 0
        if 1 <= level <= 6:
            return level

    if eng == "h1":
        return 1
    if eng == "h2":
        return 2
    if eng == "h3":
        return 3
    if eng == "h4":
        return 4

    if "제 장" in name or "큰제목" in name:
        return 1
    if "제  절" in name or "제목 2" in name:
        return 2
    if "제목 3" in name or "작은제목" in name:
        return 3
    if "제목 4" in name:
        return 4

    return None


def _is_caption(style: Optional[StyleInfo]) -> bool:
    return bool(style and "캡션" in style.name)


def _bin_data_map(zf: zipfile.ZipFile) -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    for name in zf.namelist():
        if not name.startswith("BinData/"):
            continue
        basename = Path(name).name
        if "." not in basename:
            continue
        stem = basename.split(".", 1)[0]
        mapping[stem] = name
    return mapping


def _extract_image(
    *,
    zf: zipfile.ZipFile,
    bin_map: Dict[str, str],
    binary_id: str,
    assets_dir: Path,
    extracted: Dict[str, str],
    max_image_px: int,
) -> Optional[str]:
    existing = extracted.get(binary_id)
    if existing:
        return existing

    arcname = bin_map.get(binary_id)
    if not arcname:
        return None

    suffix = Path(arcname).suffix.lower()
    if suffix == ".tmp":
        return None

    raw = zf.read(arcname)

    if suffix in {".png", ".jpg", ".jpeg"}:
        out_name = f"{binary_id}{suffix}"
        out_path = assets_dir / out_name
        out_path.write_bytes(raw)
        rel = f"assets/{out_name}"
        extracted[binary_id] = rel
        return rel

    if suffix == ".bmp":
        out_name = f"{binary_id}.jpg"
        out_path = assets_dir / out_name
        with Image.open(io.BytesIO(raw)) as image:
            image = image.convert("RGB")
            if max(image.size) > max_image_px:
                image.thumbnail((max_image_px, max_image_px))
            image.save(out_path, format="JPEG", quality=90, optimize=True)
        rel = f"assets/{out_name}"
        extracted[binary_id] = rel
        return rel

    out_name = f"{binary_id}{suffix}"
    out_path = assets_dir / out_name
    out_path.write_bytes(raw)
    rel = f"assets/{out_name}"
    extracted[binary_id] = rel
    return rel


def _first_img_ref(el: etree._Element) -> Optional[str]:
    refs = el.xpath('.//*[local-name()="img"]/@binaryItemIDRef')
    return refs[0] if refs else None


def _table_to_html(
    *,
    tbl: etree._Element,
    zf: zipfile.ZipFile,
    bin_map: Dict[str, str],
    assets_dir: Path,
    extracted: Dict[str, str],
    max_image_px: int,
) -> str:
    rows_html: List[str] = []
    for tr in tbl.xpath('./*[local-name()="tr"]'):
        cells_html: List[str] = []
        for tc in tr.xpath('./*[local-name()="tc"]'):
            parts: List[str] = []
            text = "".join(tc.xpath('.//*[local-name()="t"]/text()')).strip()
            if text:
                parts.append(html.escape(text))
            for ref in tc.xpath('.//*[local-name()="img"]/@binaryItemIDRef'):
                rel = _extract_image(
                    zf=zf,
                    bin_map=bin_map,
                    binary_id=ref,
                    assets_dir=assets_dir,
                    extracted=extracted,
                    max_image_px=max_image_px,
                )
                if rel:
                    parts.append(f'<br/><img src="{html.escape(rel)}" />')
                else:
                    parts.append(f"<br/>[이미지: {html.escape(ref)}]")
            cells_html.append(f"<td>{''.join(parts)}</td>")
        rows_html.append(f"<tr>{''.join(cells_html)}</tr>")
    return f"<table><tbody>{''.join(rows_html)}</tbody></table>"


def _paragraph_to_blocks(
    *,
    p: etree._Element,
    styles: Dict[str, StyleInfo],
    zf: zipfile.ZipFile,
    bin_map: Dict[str, str],
    assets_dir: Path,
    extracted: Dict[str, str],
    max_image_px: int,
) -> List[str]:
    style_id = p.get("styleIDRef") or ""
    style = styles.get(style_id)
    heading_level = _heading_level(style)
    is_caption = _is_caption(style)

    blocks: List[str] = []
    buffer: List[str] = []

    def flush_text() -> None:
        text = "".join(buffer).strip()
        buffer.clear()
        if not text:
            return
        if heading_level:
            blocks.append(f"<h{heading_level}>{text}</h{heading_level}>")
            return
        if is_caption:
            blocks.append(f"<p><em>{text}</em></p>")
            return
        blocks.append(f"<p>{text}</p>")

    for run in p.xpath('./*[local-name()="run"]'):
        for child in run:
            kind = _local_name(child.tag)
            if kind == "t":
                if child.text:
                    buffer.append(html.escape(child.text))
                continue
            if kind == "lineBreak":
                buffer.append("<br/>")
                continue
            if kind == "tab":
                buffer.append("\t")
                continue
            if kind == "equation":
                script = "".join(child.xpath('.//*[local-name()="script"]/text()')).strip()
                buffer.append(html.escape(script) if script else "[수식]")
                continue
            if kind == "pic":
                flush_text()
                ref = _first_img_ref(child)
                if ref:
                    rel = _extract_image(
                        zf=zf,
                        bin_map=bin_map,
                        binary_id=ref,
                        assets_dir=assets_dir,
                        extracted=extracted,
                        max_image_px=max_image_px,
                    )
                    if rel:
                        blocks.append(f'<p><img src="{html.escape(rel)}" /></p>')
                    else:
                        blocks.append(f"<p>[이미지: {html.escape(ref)}]</p>")
                continue
            if kind == "tbl":
                flush_text()
                blocks.append(
                    _table_to_html(
                        tbl=child,
                        zf=zf,
                        bin_map=bin_map,
                        assets_dir=assets_dir,
                        extracted=extracted,
                        max_image_px=max_image_px,
                    )
                )
                continue

    flush_text()
    return blocks


def convert_hwpx_to_docx(
    *,
    hwpx_path: Path,
    output_docx: Path,
    max_image_px: int,
    keep_intermediate: bool,
) -> None:
    hwpx_path = hwpx_path.resolve()
    output_docx = output_docx.resolve()
    output_docx.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(hwpx_path) as zf:
        styles = _load_styles(zf.read("Contents/header.xml"))
        section_xml = zf.read("Contents/section0.xml")
        bin_map = _bin_data_map(zf)

        root = etree.fromstring(section_xml)  # noqa: S320 (trusted local file)

        workdir_ctx = tempfile.TemporaryDirectory(prefix="hwpx2docx_")
        try:
            workdir = Path(workdir_ctx.name)
            assets_dir = workdir / "assets"
            assets_dir.mkdir(parents=True, exist_ok=True)

            extracted: Dict[str, str] = {}
            html_blocks: List[str] = [
                "<!doctype html>",
                "<html>",
                "<head>",
                '<meta charset="utf-8" />',
                "<title>Converted HWPX</title>",
                "<style>table{border-collapse:collapse}td,th{border:1px solid #999;padding:4px;vertical-align:top}</style>",
                "</head>",
                "<body>",
            ]

            for child in root:
                if _local_name(child.tag) != "p":
                    continue
                html_blocks.extend(
                    _paragraph_to_blocks(
                        p=child,
                        styles=styles,
                        zf=zf,
                        bin_map=bin_map,
                        assets_dir=assets_dir,
                        extracted=extracted,
                        max_image_px=max_image_px,
                    )
                )

            html_blocks.extend(["</body>", "</html>"])
            html_path = workdir / "converted.html"
            html_path.write_text("\n".join(html_blocks), encoding="utf-8")

            pandoc = shutil.which("pandoc")
            if not pandoc:
                raise RuntimeError("pandoc is required but was not found in PATH.")

            subprocess.run(
                [pandoc, str(html_path), "-f", "html", "-t", "docx", "-o", str(output_docx)],
                check=True,
                cwd=workdir,
            )
            if keep_intermediate:
                intermediate_dir = output_docx.with_suffix(".hwpx2docx")
                if intermediate_dir.exists():
                    shutil.rmtree(intermediate_dir)
                shutil.copytree(workdir, intermediate_dir)
        finally:
            workdir_ctx.cleanup()


def main() -> int:
    parser = argparse.ArgumentParser(description="Best-effort HWPX → DOCX converter (text/tables/images).")
    parser.add_argument("hwpx", type=Path, help="Input .hwpx file")
    parser.add_argument("--output", type=Path, required=True, help="Output .docx path")
    parser.add_argument(
        "--max-image-px",
        type=int,
        default=2200,
        help="Downscale embedded BMP images so their max dimension is <= this value (default: 2200).",
    )
    parser.add_argument(
        "--keep-intermediate",
        action="store_true",
        help="Keep extracted assets + intermediate HTML next to the output (suffix: .hwpx2docx).",
    )
    args = parser.parse_args()

    convert_hwpx_to_docx(
        hwpx_path=args.hwpx,
        output_docx=args.output,
        max_image_px=args.max_image_px,
        keep_intermediate=args.keep_intermediate,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
