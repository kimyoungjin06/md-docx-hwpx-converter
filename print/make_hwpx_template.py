#!/usr/bin/env python3
from __future__ import annotations

import argparse
import zipfile
from pathlib import Path
from typing import Optional

from lxml import etree


def _local_name(tag: str) -> str:
    if "}" in tag:
        return tag.rsplit("}", 1)[-1]
    return tag


def _minify_section0(section_xml: bytes) -> bytes:
    root = etree.fromstring(section_xml)  # noqa: S320 (trusted local file)

    sec_pr_p: Optional[etree._Element] = None
    for child in root:
        if _local_name(child.tag) != "p":
            continue
        if child.xpath('.//*[local-name()="secPr"]'):
            sec_pr_p = child
            break

    if sec_pr_p is None:
        raise RuntimeError("Could not find <hp:secPr> in Contents/section0.xml (template input not supported).")

    for child in list(root):
        if child is sec_pr_p:
            continue
        root.remove(child)

    return etree.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=True,
        standalone=True,
    )


def _minify_content_hpf(content_hpf: bytes) -> bytes:
    root = etree.fromstring(content_hpf)  # noqa: S320 (trusted local file)

    manifest_nodes = root.xpath('//*[local-name()="manifest"]')
    if manifest_nodes:
        manifest = manifest_nodes[0]
        for item in list(manifest.xpath('./*[local-name()="item"]')):
            href = (item.get("href") or "").strip()
            if href.startswith("BinData/"):
                manifest.remove(item)

        allowed_ids = {el.get("id") for el in manifest.xpath('./*[local-name()="item"]') if el.get("id")}
        for spine in root.xpath('//*[local-name()="spine"]'):
            for itemref in list(spine.xpath('./*[local-name()="itemref"]')):
                idref = (itemref.get("idref") or "").strip()
                if idref and idref not in allowed_ids:
                    spine.remove(itemref)

    return etree.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=True,
        standalone=True,
    )


def make_hwpx_template(*, input_hwpx: Path, output_hwpx: Path) -> None:
    input_hwpx = input_hwpx.resolve()
    output_hwpx = output_hwpx.resolve()
    output_hwpx.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(input_hwpx) as zf_in:
        section0_xml = _minify_section0(zf_in.read("Contents/section0.xml"))
        content_hpf = _minify_content_hpf(zf_in.read("Contents/content.hpf"))

        with zipfile.ZipFile(output_hwpx, "w") as zf_out:
            for info in zf_in.infolist():
                if info.filename.startswith("BinData/"):
                    continue
                data = zf_in.read(info.filename)
                if info.filename == "Contents/section0.xml":
                    data = section0_xml
                elif info.filename == "Contents/content.hpf":
                    data = content_hpf
                zf_out.writestr(info, data)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Create a lightweight HWPX template by stripping content + BinData from an existing .hwpx."
    )
    parser.add_argument("input_hwpx", type=Path, help="Source .hwpx (any file with correct styles/settings)")
    parser.add_argument("--output", type=Path, required=True, help="Output template .hwpx path")
    args = parser.parse_args()

    make_hwpx_template(input_hwpx=args.input_hwpx, output_hwpx=args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

