#!/usr/bin/env python
from __future__ import annotations

import argparse
import json
import posixpath
import re
import sys
import zipfile
from pathlib import Path
from typing import Any
import xml.etree.ElementTree as ET


RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
PML_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PLACEHOLDER_RE = re.compile(r"\{\{[^{}]+\}\}")
REQUIRED_PARTS = (
    "[Content_Types].xml",
    "_rels/.rels",
    "ppt/presentation.xml",
    "ppt/_rels/presentation.xml.rels",
)


def add_issue(
    issues: list[dict[str, Any]],
    severity: str,
    category: str,
    message: str,
    *,
    slide: int | None = None,
    part: str | None = None,
    details: dict[str, Any] | None = None,
) -> None:
    issue: dict[str, Any] = {"severity": severity, "category": category, "message": message}
    if slide is not None:
        issue["slide"] = slide
    if part is not None:
        issue["part"] = part
    if details:
        issue["details"] = details
    issues.append(issue)


def finalize(report: dict[str, Any]) -> dict[str, Any]:
    counts = {"error": 0, "warning": 0, "info": 0}
    for issue in report["issues"]:
        severity = issue.get("severity", "info")
        counts[severity] = counts.get(severity, 0) + 1
    report.setdefault("summary", {})["issues"] = counts
    if counts["error"]:
        report["status"] = "error"
    elif counts["warning"]:
        report["status"] = "warning"
    else:
        report["status"] = "pass"
    return report


def read_xml(package: zipfile.ZipFile, part: str) -> ET.Element:
    return ET.fromstring(package.read(part))


def media_parts(parts: set[str]) -> list[str]:
    return sorted(part for part in parts if part.startswith("ppt/media/") and not part.endswith("/"))


def natural_key(part: str) -> tuple[Any, ...]:
    return tuple(int(chunk) if chunk.isdigit() else chunk for chunk in re.split(r"(\d+)", part))


def rels_owner(rels_part: str) -> str:
    if rels_part == "_rels/.rels":
        return ""
    if "/_rels/" in rels_part:
        base, tail = rels_part.split("/_rels/", 1)
        return posixpath.join(base, tail[:-5])
    return rels_part[:-5]


def resolve_target(rels_part: str, target: str) -> str:
    if not target:
        return ""
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    owner = rels_owner(rels_part)
    base = posixpath.dirname(owner) if owner else ""
    return posixpath.normpath(posixpath.join(base, target))


def collect_rels(package: zipfile.ZipFile, issues: list[dict[str, Any]]) -> dict[str, list[dict[str, Any]]]:
    rels_by_part: dict[str, list[dict[str, Any]]] = {}
    for rels_part in package.namelist():
        if not rels_part.endswith(".rels"):
            continue
        try:
            root = read_xml(package, rels_part)
        except ET.ParseError as exc:
            add_issue(issues, "error", "relationships", f"Relationship XML parse error: {exc}", part=rels_part)
            continue
        rels = []
        for rel in root.findall(f"{{{RELS_NS}}}Relationship"):
            mode = rel.attrib.get("TargetMode", "")
            target = rel.attrib.get("Target", "")
            rels.append({
                "id": rel.attrib.get("Id", ""),
                "type": rel.attrib.get("Type", ""),
                "target": target,
                "target_mode": mode,
                "resolved": None if mode.lower() == "external" else resolve_target(rels_part, target),
            })
        rels_by_part[rels_part] = rels
    return rels_by_part


def presentation_slide_parts(package: zipfile.ZipFile, rels_by_part: dict[str, list[dict[str, Any]]], issues: list[dict[str, Any]]) -> list[str]:
    root = read_xml(package, "ppt/presentation.xml")
    rels = {rel["id"]: rel for rel in rels_by_part.get("ppt/_rels/presentation.xml.rels", [])}
    slide_parts: list[str] = []
    for slide_id in root.findall(f".//{{{PML_NS}}}sldId"):
        rid = slide_id.attrib.get(f"{{{OFFICE_REL_NS}}}id")
        rel = rels.get(rid or "")
        if not rel:
            add_issue(
                issues,
                "error",
                "relationships",
                "Presentation slide entry references a missing relationship id.",
                part="ppt/presentation.xml",
                details={"relationship_id": rid},
            )
            continue
        if rel.get("resolved"):
            slide_parts.append(str(rel["resolved"]))
    if not slide_parts:
        add_issue(issues, "error", "structure", "Presentation does not contain any slides.")
    return slide_parts


def slide_size(package: zipfile.ZipFile, issues: list[dict[str, Any]]) -> tuple[int, int]:
    try:
        root = read_xml(package, "ppt/presentation.xml")
        size = root.find(f"{{{PML_NS}}}sldSz")
        if size is not None:
            return int(size.attrib.get("cx", "9144000")), int(size.attrib.get("cy", "5143500"))
    except Exception:
        pass
    add_issue(issues, "warning", "structure", "Missing or malformed slide size; assuming 16:9 widescreen.")
    return 9144000, 5143500


def referenced_parts(rels_by_part: dict[str, list[dict[str, Any]]]) -> set[str]:
    refs = set()
    for rels in rels_by_part.values():
        for rel in rels:
            if rel.get("resolved"):
                refs.add(str(rel["resolved"]))
    return refs


def check_relationship_targets(names: set[str], rels_by_part: dict[str, list[dict[str, Any]]], issues: list[dict[str, Any]]) -> None:
    for rels_part, rels in rels_by_part.items():
        for rel in rels:
            if str(rel.get("target_mode", "")).lower() == "external":
                continue
            target = rel.get("resolved")
            if target and target not in names:
                add_issue(
                    issues,
                    "error",
                    "relationships",
                    "Relationship target is missing from the PPTX package.",
                    part=rels_part,
                    details={"relationship_id": rel.get("id"), "target": rel.get("target"), "resolved": target},
                )


def shape_name(shape: ET.Element) -> str:
    c_nv_pr = shape.find(f".//{{{PML_NS}}}cNvPr")
    return c_nv_pr.attrib.get("name", "") if c_nv_pr is not None else ""


def check_shape_bounds(root: ET.Element, size: tuple[int, int], slide_index: int, slide_part: str, issues: list[dict[str, Any]]) -> None:
    slide_cx, slide_cy = size
    shape_tags = {f"{{{PML_NS}}}sp", f"{{{PML_NS}}}pic", f"{{{PML_NS}}}graphicFrame"}
    for shape in root.iter():
        if shape.tag not in shape_tags:
            continue
        xfrm = shape.find(f".//{{{DML_NS}}}xfrm")
        if xfrm is None:
            continue
        off = xfrm.find(f"{{{DML_NS}}}off")
        ext = xfrm.find(f"{{{DML_NS}}}ext")
        if off is None or ext is None:
            continue
        try:
            x = int(float(off.attrib.get("x", "0")))
            y = int(float(off.attrib.get("y", "0")))
            cx = int(float(ext.attrib.get("cx", "0")))
            cy = int(float(ext.attrib.get("cy", "0")))
        except ValueError:
            add_issue(issues, "warning", "shapes", "Shape transform contains non-numeric bounds.", slide=slide_index, part=slide_part)
            continue
        if cx < 0 or cy < 0:
            add_issue(issues, "error", "shapes", "Shape has negative size.", slide=slide_index, part=slide_part)
        if x + cx < 0 or y + cy < 0 or x > slide_cx or y > slide_cy:
            add_issue(
                issues,
                "warning",
                "shapes",
                "Shape is entirely outside the slide canvas.",
                slide=slide_index,
                part=slide_part,
                details={"shape": shape_name(shape), "x": x, "y": y, "cx": cx, "cy": cy},
            )


def slide_rels_path(slide_part: str) -> str:
    directory, filename = posixpath.split(slide_part)
    return posixpath.join(directory, "_rels", f"{filename}.rels")


def slide_images(root: ET.Element, slide_part: str, rels_by_part: dict[str, list[dict[str, Any]]]) -> list[dict[str, Any]]:
    rels = {rel["id"]: rel for rel in rels_by_part.get(slide_rels_path(slide_part), [])}
    images = []
    for blip in root.iter(f"{{{DML_NS}}}blip"):
        rid = blip.attrib.get(f"{{{OFFICE_REL_NS}}}embed") or blip.attrib.get(f"{{{OFFICE_REL_NS}}}link")
        if not rid:
            continue
        rel = rels.get(rid)
        images.append({
            "relationship_id": rid,
            "part": str(rel.get("resolved")) if rel else "",
            "target": rel.get("target") if rel else "",
            "missing_relationship": rel is None,
        })
    return images


def check_slide_text(text: str, index: int, part: str, issues: list[dict[str, Any]]) -> None:
    if not text.strip():
        add_issue(issues, "warning", "content", "Slide contains no extractable text.", slide=index, part=part)
    if PLACEHOLDER_RE.search(text):
        add_issue(issues, "error", "content", "Slide contains unresolved template placeholder text.", slide=index, part=part)
    if re.search(r"(?im)^\s*layout\s*:", text):
        add_issue(issues, "error", "content", "Slide still contains a layout marker.", slide=index, part=part)
    if "????" in text:
        add_issue(issues, "error", "content", "Slide contains question-mark mojibake.", slide=index, part=part)
    if len(text) > 1200:
        add_issue(issues, "warning", "content", "Slide has a large amount of extractable text.", slide=index, part=part, details={"text_length": len(text)})


def inspect_slide(
    package: zipfile.ZipFile,
    names: set[str],
    part: str,
    index: int,
    rels_by_part: dict[str, list[dict[str, Any]]],
    size: tuple[int, int],
    issues: list[dict[str, Any]],
) -> dict[str, Any]:
    if part not in names:
        add_issue(issues, "error", "structure", "Slide part is missing.", slide=index, part=part)
        return {"index": index, "part": part, "text": "", "source_captions": [], "image_count": 0}
    root = read_xml(package, part)
    texts = [node.text or "" for node in root.iter(f"{{{DML_NS}}}t") if (node.text or "").strip()]
    text = "\n".join(texts)
    check_slide_text(text, index, part, issues)
    check_shape_bounds(root, size, index, part, issues)
    images = slide_images(root, part, rels_by_part)
    for image in images:
        if image["missing_relationship"] or image["part"] not in names:
            add_issue(issues, "error", "resources", "Slide picture references a missing media part.", slide=index, part=part, details=image)
    return {
        "index": index,
        "part": part,
        "text": text,
        "text_length": len(text),
        "source_captions": [line for item in texts for line in item.splitlines() if "Source:" in line],
        "image_count": len(images),
        "images": images,
    }


def image_dimensions(data: bytes) -> tuple[int | None, int | None]:
    if data.startswith(b"\x89PNG\r\n\x1a\n") and len(data) >= 24:
        return int.from_bytes(data[16:20], "big"), int.from_bytes(data[20:24], "big")
    if data.startswith(b"\xff\xd8"):
        idx = 2
        while idx + 9 < len(data):
            if data[idx] != 0xFF:
                idx += 1
                continue
            marker = data[idx + 1]
            idx += 2
            if marker in {0xD8, 0xD9}:
                continue
            length = int.from_bytes(data[idx:idx + 2], "big")
            if length < 2 or idx + length > len(data):
                break
            if 0xC0 <= marker <= 0xC3 or 0xC5 <= marker <= 0xC7 or 0xC9 <= marker <= 0xCB or 0xCD <= marker <= 0xCF:
                return int.from_bytes(data[idx + 5:idx + 7], "big"), int.from_bytes(data[idx + 3:idx + 5], "big")
            idx += length
    return None, None


def qa_pptx(path: Path, expected_slides: int | None = None, required_sources: list[str] | None = None) -> dict[str, Any]:
    issues: list[dict[str, Any]] = []
    report: dict[str, Any] = {
        "path": str(path.resolve()),
        "exists": path.exists(),
        "status": "error",
        "summary": {},
        "slides": [],
        "media": [],
        "issues": issues,
    }
    if not path.exists():
        add_issue(issues, "error", "package", "File does not exist.", part=str(path))
        return finalize(report)
    try:
        with zipfile.ZipFile(path) as package:
            names = set(package.namelist())
            bad_member = package.testzip()
            if bad_member:
                add_issue(issues, "error", "package", f"Zip member failed CRC check: {bad_member}", part=bad_member)
            for part in REQUIRED_PARTS:
                if part not in names:
                    add_issue(issues, "error", "package", "Required PPTX part is missing.", part=part)
            if "ppt/presentation.xml" not in names:
                return finalize(report)
            rels_by_part = collect_rels(package, issues)
            slides = presentation_slide_parts(package, rels_by_part, issues)
            direct_slides = sorted((name for name in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)), key=natural_key)
            if direct_slides and len(direct_slides) != len(slides):
                add_issue(issues, "warning", "structure", "Slide XML part count differs from presentation slide list.")
            if expected_slides is not None and len(slides) != expected_slides:
                add_issue(issues, "error", "structure", "Slide count does not match expected value.", details={"expected": expected_slides, "actual": len(slides)})
            refs = referenced_parts(rels_by_part)
            check_relationship_targets(names, rels_by_part, issues)
            for media in media_parts(names):
                if media not in refs:
                    add_issue(issues, "warning", "resources", "Media file is present but not referenced by any relationship.", part=media)
            size = slide_size(package, issues)
            all_text = []
            for index, slide_part in enumerate(slides or direct_slides, start=1):
                slide_report = inspect_slide(package, names, slide_part, index, rels_by_part, size, issues)
                report["slides"].append(slide_report)
                all_text.append(slide_report.get("text", ""))
            full_text = "\n".join(all_text)
            for source in required_sources or []:
                if source and source not in full_text:
                    add_issue(issues, "error", "content", "Required source caption/text was not found.", details={"required_source": source})
            for media in media_parts(names):
                data = package.read(media)
                width, height = image_dimensions(data)
                report["media"].append({"part": media, "size_bytes": len(data), "referenced": media in refs, "width_px": width, "height_px": height})
            report["summary"] = {
                "slide_count": len(slides or direct_slides),
                "media_count": len(media_parts(names)),
                "referenced_media_count": len(media_parts(refs)),
                "relationship_count": sum(len(rels) for rels in rels_by_part.values()),
                "slide_width_emu": size[0],
                "slide_height_emu": size[1],
            }
    except zipfile.BadZipFile:
        add_issue(issues, "error", "package", "File is not a valid zip/PPTX package.", part=str(path))
    except ET.ParseError as exc:
        add_issue(issues, "error", "xml", f"XML parse error: {exc}")
    return finalize(report)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run bottom-layer PPTX package/XML QA.")
    parser.add_argument("pptx", type=Path, help="PPTX file to inspect.")
    parser.add_argument("--expected-slides", type=int, help="Expected slide count.")
    parser.add_argument("--require-source", action="append", default=[], help="Text/source caption that must exist.")
    parser.add_argument("--report", type=Path, help="Write the full JSON report.")
    parser.add_argument("--strict", action="store_true", help="Exit non-zero unless status is pass.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    report = qa_pptx(args.pptx, args.expected_slides, args.require_source)
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({
        "status": report["status"],
        "summary": report.get("summary", {}),
        "issues": report.get("issues", []),
        "report": str(args.report) if args.report else None,
    }, ensure_ascii=False, indent=2))
    return 1 if args.strict and report["status"] != "pass" else 0


if __name__ == "__main__":
    raise SystemExit(main())
