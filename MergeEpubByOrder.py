#!/usr/bin/env python3
"""Merge chapter EPUBs into one EPUB2 anthology ordered by provided plan."""

from __future__ import annotations

import argparse
import json
import os
import posixpath
import re
import sys
import uuid
import zipfile
from dataclasses import dataclass
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Dict, List, Optional, Tuple
import xml.etree.ElementTree as ET

OPF_NS = "http://www.idpf.org/2007/opf"
DC_NS = "http://purl.org/dc/elements/1.1/"
NCX_NS = "http://www.daisy.org/z3986/2005/ncx/"
CONT_NS = "urn:oasis:names:tc:opendocument:xmlns:container"

ET.register_namespace("", OPF_NS)
ET.register_namespace("dc", DC_NS)
ET.register_namespace("opf", OPF_NS)


class MergeError(RuntimeError):
    pass


@dataclass
class ChapterPlan:
    chapter_name: str
    order: Optional[float]
    epub_path: str


@dataclass
class MergePlan:
    output_epub_path: str
    title: str
    author: str
    language: str
    description: str
    contributor: str
    chapters: List[ChapterPlan]


def _ns(tag: str, ns: str) -> str:
    """Build a namespaced XML tag string."""
    return f"{{{ns}}}{tag}"


def _local(tag: str) -> str:
    """Extract local name from an XML tag."""
    return tag.split("}", 1)[1] if "}" in tag else tag


def _read_text(zf: zipfile.ZipFile, name: str) -> str:
    """Read a UTF-8 text entry from a zip, raising MergeError when missing."""
    try:
        return zf.read(name).decode("utf-8")
    except KeyError as exc:
        raise MergeError(f"missing entry in epub: {name}") from exc


def _find_rootfile_path(container_xml: str) -> str:
    """Resolve OPF path from EPUB container.xml."""
    root = ET.fromstring(container_xml)
    rootfile = root.find(f".//{{{CONT_NS}}}rootfile")
    if rootfile is None:
        rootfile = root.find(".//rootfile")
    if rootfile is None:
        raise MergeError("container.xml missing rootfile")
    full_path = rootfile.attrib.get("full-path", "").strip()
    if not full_path:
        raise MergeError("container.xml rootfile missing full-path")
    return full_path


def _norm(path: str) -> str:
    """Normalize to stable POSIX-style relative path."""
    return posixpath.normpath(path.replace("\\", "/"))


def _safe_id(raw: str) -> str:
    """Sanitize id values so merged OPF ids stay XML-safe."""
    return re.sub(r"[^A-Za-z0-9_.-]", "_", raw) or "id"


def _first_existing_href(old_to_new_href: Dict[str, str], keys: List[str]) -> Optional[str]:
    """Return the first mapped href from candidate keys."""
    for key in keys:
        if key in old_to_new_href:
            return old_to_new_href[key]
    return None


def _load_plan(path: str) -> MergePlan:
    # Plan json is produced by the PowerShell pipeline and defines
    # both ordering and metadata for the merged anthology.
    with open(path, "r", encoding="utf-8") as fp:
        raw = json.load(fp)

    required = ["output_epub_path", "title", "author", "language", "description", "chapters"]
    for key in required:
        if key not in raw:
            raise MergeError(f"plan json missing field: {key}")

    chapters = []
    for idx, item in enumerate(raw["chapters"]):
        for key in ["chapter_name", "epub_path"]:
            if key not in item:
                raise MergeError(f"plan chapter[{idx}] missing field: {key}")
        order = item.get("order")
        order_value = float(order) if order is not None else None
        chapters.append(
            ChapterPlan(
                chapter_name=str(item["chapter_name"]),
                order=order_value,
                epub_path=str(item["epub_path"]),
            )
        )

    if not chapters:
        raise MergeError("no chapters provided")

    return MergePlan(
        output_epub_path=str(raw["output_epub_path"]),
        title=str(raw["title"]),
        author=str(raw["author"]),
        language=str(raw["language"]),
        description=str(raw["description"]),
        contributor=str(raw.get("contributor", "mjnai-merge")),
        chapters=chapters,
    )


def _write_container_xml(out: zipfile.ZipFile) -> None:
    """Write EPUB META-INF/container.xml that points to merged content.opf."""
    container = ET.Element(_ns("container", CONT_NS), attrib={"version": "1.0"})
    rootfiles = ET.SubElement(container, _ns("rootfiles", CONT_NS))
    ET.SubElement(
        rootfiles,
        _ns("rootfile", CONT_NS),
        attrib={
            "full-path": "content.opf",
            "media-type": "application/oebps-package+xml",
        },
    )
    xml = ET.tostring(container, encoding="utf-8", xml_declaration=True)
    out.writestr("META-INF/container.xml", xml, compress_type=zipfile.ZIP_DEFLATED)


def _write_toc_ncx(
    out: zipfile.ZipFile,
    uid: str,
    title: str,
    navpoints: List[Tuple[str, str]],
) -> None:
    """Create a flat NCX table of contents from chapter navpoints."""
    ncx = ET.Element(_ns("ncx", NCX_NS), attrib={"version": "2005-1"})
    head = ET.SubElement(ncx, _ns("head", NCX_NS))
    ET.SubElement(head, _ns("meta", NCX_NS), attrib={"name": "dtb:uid", "content": uid})
    ET.SubElement(head, _ns("meta", NCX_NS), attrib={"name": "dtb:depth", "content": "1"})
    ET.SubElement(head, _ns("meta", NCX_NS), attrib={"name": "dtb:totalPageCount", "content": "0"})
    ET.SubElement(head, _ns("meta", NCX_NS), attrib={"name": "dtb:maxPageNumber", "content": "0"})

    doc_title = ET.SubElement(ncx, _ns("docTitle", NCX_NS))
    ET.SubElement(doc_title, _ns("text", NCX_NS)).text = title

    nav_map = ET.SubElement(ncx, _ns("navMap", NCX_NS))
    for idx, (name, href) in enumerate(navpoints, start=1):
        nav = ET.SubElement(nav_map, _ns("navPoint", NCX_NS), attrib={"id": f"book{idx:03d}", "playOrder": str(idx)})
        nav_label = ET.SubElement(nav, _ns("navLabel", NCX_NS))
        ET.SubElement(nav_label, _ns("text", NCX_NS)).text = name
        ET.SubElement(nav, _ns("content", NCX_NS), attrib={"src": href})

    xml = ET.tostring(ncx, encoding="utf-8", xml_declaration=True)
    out.writestr("toc.ncx", xml, compress_type=zipfile.ZIP_DEFLATED)


def _write_content_opf(
    out: zipfile.ZipFile,
    uid: str,
    plan: MergePlan,
    manifest_items: List[Tuple[str, str, str]],
    spine_ids: List[str],
    cover_image_href: Optional[str],
) -> None:
    """Write merged OPF metadata, manifest, spine, and optional cover guide."""
    package = ET.Element(
        _ns("package", OPF_NS),
        attrib={
            "version": "2.0",
            "unique-identifier": "mjnai-merge-id",
        },
    )

    metadata = ET.SubElement(
        package,
        _ns("metadata", OPF_NS),
        attrib={
            _ns("dc", "http://www.w3.org/2000/xmlns/"): DC_NS,
            _ns("opf", "http://www.w3.org/2000/xmlns/"): OPF_NS,
        },
    )
    ET.SubElement(metadata, _ns("identifier", DC_NS), attrib={"id": "mjnai-merge-id"}).text = uid
    ET.SubElement(metadata, _ns("title", DC_NS)).text = plan.title
    ET.SubElement(
        metadata,
        _ns("creator", DC_NS),
        attrib={_ns("role", OPF_NS): "aut", _ns("file-as", OPF_NS): plan.author},
    ).text = plan.author
    ET.SubElement(metadata, _ns("language", DC_NS)).text = plan.language
    ET.SubElement(metadata, _ns("description", DC_NS)).text = plan.description
    ET.SubElement(metadata, _ns("contributor", DC_NS)).text = plan.contributor

    if cover_image_href:
        ET.SubElement(metadata, _ns("meta", OPF_NS), attrib={"name": "cover", "content": "coverimageid"})

    manifest = ET.SubElement(package, _ns("manifest", OPF_NS))
    if cover_image_href:
        media = "image/jpeg"
        ext = Path(cover_image_href).suffix.lower()
        if ext == ".png":
            media = "image/png"
        elif ext == ".gif":
            media = "image/gif"
        elif ext == ".webp":
            media = "image/webp"
        ET.SubElement(manifest, _ns("item", OPF_NS), attrib={"id": "coverimageid", "href": cover_image_href, "media-type": media})
        ET.SubElement(manifest, _ns("item", OPF_NS), attrib={"id": "cover", "href": "cover.xhtml", "media-type": "application/xhtml+xml"})
    ET.SubElement(manifest, _ns("item", OPF_NS), attrib={"id": "ncx", "href": "toc.ncx", "media-type": "application/x-dtbncx+xml"})

    for item_id, href, media_type in manifest_items:
        ET.SubElement(manifest, _ns("item", OPF_NS), attrib={"id": item_id, "href": href, "media-type": media_type})

    spine = ET.SubElement(package, _ns("spine", OPF_NS), attrib={"toc": "ncx"})
    if cover_image_href:
        ET.SubElement(spine, _ns("itemref", OPF_NS), attrib={"idref": "cover", "linear": "yes"})
    for sid in spine_ids:
        ET.SubElement(spine, _ns("itemref", OPF_NS), attrib={"idref": sid, "linear": "yes"})

    if cover_image_href:
        guide = ET.SubElement(package, _ns("guide", OPF_NS))
        ET.SubElement(guide, _ns("reference", OPF_NS), attrib={"type": "cover", "title": "Cover", "href": "cover.xhtml"})

    xml = ET.tostring(package, encoding="utf-8", xml_declaration=True)
    out.writestr("content.opf", xml, compress_type=zipfile.ZIP_DEFLATED)


def _write_cover_xhtml(out: zipfile.ZipFile, cover_href: str) -> None:
    """Write a simple XHTML cover page referencing the selected cover image."""
    html = f"""<?xml version=\"1.0\" encoding=\"utf-8\"?>
<html xmlns=\"http://www.w3.org/1999/xhtml\">
  <head><title>Cover</title></head>
  <body>
    <div style=\"text-align:center;\"><img alt=\"cover\" src=\"{cover_href}\" /></div>
  </body>
</html>
"""
    out.writestr("cover.xhtml", html.encode("utf-8"), compress_type=zipfile.ZIP_DEFLATED)


def merge(plan: MergePlan) -> None:
    """Merge chapter EPUBs into one EPUB2 archive based on plan order."""
    out_path = Path(plan.output_epub_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    tmp_file = NamedTemporaryFile(prefix="mjnai_merge_", suffix=".epub", delete=False, dir=str(out_path.parent))
    tmp_path = Path(tmp_file.name)
    tmp_file.close()

    manifest_items: List[Tuple[str, str, str]] = []
    spine_ids: List[str] = []
    navpoints: List[Tuple[str, str]] = []
    seen_hrefs = set()

    cover_href: Optional[str] = None
    uid = f"urn:uuid:{uuid.uuid4()}"

    try:
        with zipfile.ZipFile(tmp_path, "w", allowZip64=True) as out:
            # EPUB requirement: mimetype must be first and uncompressed.
            mime_info = zipfile.ZipInfo("mimetype")
            mime_info.compress_type = zipfile.ZIP_STORED
            out.writestr(mime_info, b"application/epub+zip")
            _write_container_xml(out)

            for chapter_index, chapter in enumerate(plan.chapters, start=1):
                epub_path = Path(chapter.epub_path)
                if not epub_path.is_file():
                    raise MergeError(f"chapter epub not found: {epub_path}")

                with zipfile.ZipFile(epub_path, "r") as src:
                    container_xml = _read_text(src, "META-INF/container.xml")
                    rootfile = _find_rootfile_path(container_xml)
                    opf_xml = _read_text(src, rootfile)

                    opf_root = ET.fromstring(opf_xml)
                    root_dir = posixpath.dirname(rootfile)

                    manifest = opf_root.find(f"{{{OPF_NS}}}manifest")
                    if manifest is None:
                        manifest = opf_root.find("manifest")
                    spine = opf_root.find(f"{{{OPF_NS}}}spine")
                    if spine is None:
                        spine = opf_root.find("spine")
                    metadata = opf_root.find(f"{{{OPF_NS}}}metadata")
                    if metadata is None:
                        metadata = opf_root.find("metadata")

                    if manifest is None or spine is None:
                        raise MergeError(f"invalid opf in {epub_path}: missing manifest/spine")

                    # Namespace each source book under "<index>/" to avoid
                    # href/id collisions between different chapter EPUBs.
                    chapter_prefix = f"{chapter_index}/"
                    old_to_new_id: Dict[str, str] = {}
                    old_to_new_href: Dict[str, str] = {}

                    source_cover_id = None
                    if metadata is not None:
                        for meta in metadata.findall(f"{{{OPF_NS}}}meta") + metadata.findall("meta"):
                            if meta.attrib.get("name") == "cover":
                                source_cover_id = meta.attrib.get("content")
                                break

                    for item in manifest:
                        if _local(item.tag) != "item":
                            continue
                        old_id = item.attrib.get("id", "")
                        href = item.attrib.get("href", "")
                        media_type = item.attrib.get("media-type", "application/octet-stream")
                        if not href:
                            continue

                        src_name = _norm(posixpath.join(root_dir, href))
                        dst_name = _norm(chapter_prefix + src_name)

                        try:
                            blob = src.read(src_name)
                        except KeyError:
                            # Keep merge resilient to occasional bad item refs.
                            continue

                        if dst_name not in seen_hrefs:
                            out.writestr(dst_name, blob, compress_type=zipfile.ZIP_DEFLATED)
                            seen_hrefs.add(dst_name)

                        new_id = f"c{chapter_index}_{_safe_id(old_id or Path(href).stem)}"
                        old_to_new_id[old_id] = new_id
                        old_to_new_href[old_id] = dst_name
                        manifest_items.append((new_id, dst_name, media_type))

                    first_spine_href: Optional[str] = None
                    for itemref in spine:
                        if _local(itemref.tag) != "itemref":
                            continue
                        old_idref = itemref.attrib.get("idref", "")
                        new_idref = old_to_new_id.get(old_idref)
                        if not new_idref:
                            continue
                        spine_ids.append(new_idref)
                        if first_spine_href is None:
                            first_spine_href = old_to_new_href.get(old_idref)

                    if first_spine_href is None:
                        raise MergeError(f"cannot resolve first spine item for {epub_path}")

                    # Each chapter contributes one top-level TOC entry.
                    navpoints.append((chapter.chapter_name, first_spine_href))

                    if cover_href is None and source_cover_id:
                        cover_href = old_to_new_href.get(source_cover_id)

            if cover_href:
                _write_cover_xhtml(out, cover_href)

            _write_toc_ncx(out, uid=uid, title=plan.title, navpoints=navpoints)
            _write_content_opf(
                out,
                uid=uid,
                plan=plan,
                manifest_items=manifest_items,
                spine_ids=spine_ids,
                cover_image_href=cover_href,
            )

        # Atomic replace prevents leaving a partial EPUB on failure.
        os.replace(tmp_path, out_path)
    except Exception:
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass
        raise


def build_arg_parser() -> argparse.ArgumentParser:
    """Define CLI interface for the merge helper."""
    p = argparse.ArgumentParser(description="Merge ordered chapter EPUBs into one EPUB2 file")
    p.add_argument("--plan", required=True, help="Path to merge plan JSON")
    return p


def main(argv: List[str]) -> int:
    """CLI entrypoint with stable exit codes for pipeline integration."""
    parser = build_arg_parser()
    args = parser.parse_args(argv)

    def _safe_text(value: object) -> str:
        text = str(value)
        enc = sys.stdout.encoding or "utf-8"
        return text.encode(enc, errors="backslashreplace").decode(enc, errors="replace")

    try:
        plan = _load_plan(args.plan)
        merge(plan)
        print(f"MERGE_OK: {_safe_text(plan.output_epub_path)}")
        return 0
    except MergeError as exc:
        print(f"MERGE_ERROR: {_safe_text(exc)}", file=sys.stderr)
        return 2
    except Exception as exc:  # pragma: no cover
        print(f"MERGE_ERROR_UNEXPECTED: {_safe_text(exc)}", file=sys.stderr)
        return 3


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
