#!/usr/bin/env python
from __future__ import annotations

import argparse
import json
import math
import os
import sys
import zlib
from pathlib import Path
from typing import Any


SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import preview_deck_qa as preview_qa  # noqa: E402


ROLE_CHECKS = {
    "cover": [
        "Title is dominant, readable, and not inside a decorative card.",
        "Subtitle/venue metadata are secondary and aligned with the template.",
        "No stray white panel or leftover sample text is visible.",
    ],
    "contents": [
        "Entries are balanced and ordered like the talk narrative.",
        "Numbers and labels do not collide with the background or each other.",
        "Contents page uses the template contents layout, not a section divider.",
    ],
    "section": [
        "Large section numeral and section title do not overlap.",
        "Dark/light template background is preserved without unexpected white blocks.",
        "Only transition-level content appears; no dense evidence text is carried here.",
    ],
    "content": [
        "One central claim is visually clear.",
        "Body text is concise, aligned, and not too close to edges.",
        "Decorative template elements remain secondary to the argument.",
    ],
    "evidence": [
        "Source screenshot/table is legible enough to support the claim.",
        "SCUT-blue highlight points to the exact evidence being discussed.",
        "Source caption is present and does not collide with content.",
    ],
    "closing": [
        "Closing message is centered and visually calm.",
        "No leftover helper text, layout marker, or blank white panel is visible.",
        "The ending matches the same template family as the cover.",
    ],
}

ROLE_COLORS = {
    "cover": (45, 79, 143, 255),
    "contents": (118, 171, 220, 255),
    "section": (31, 55, 107, 255),
    "content": (183, 201, 232, 255),
    "evidence": (45, 79, 143, 255),
    "closing": (30, 30, 30, 255),
}

DIGIT_FONT = {
    "0": ("111", "101", "101", "101", "111"),
    "1": ("010", "110", "010", "010", "111"),
    "2": ("111", "001", "111", "100", "111"),
    "3": ("111", "001", "111", "001", "111"),
    "4": ("101", "101", "111", "001", "001"),
    "5": ("111", "100", "111", "001", "111"),
    "6": ("111", "100", "111", "101", "111"),
    "7": ("111", "001", "001", "001", "001"),
    "8": ("111", "101", "111", "101", "111"),
    "9": ("111", "101", "111", "001", "111"),
}


def parse_slide_roles(values: list[str]) -> dict[int, str]:
    roles: dict[int, str] = {}
    for value in values:
        if ":" not in value:
            raise ValueError(f"Invalid --slide-role value: {value!r}; use INDEX:role.")
        index_text, role = value.split(":", 1)
        index = int(index_text.strip())
        role = normalize_role(role)
        roles[index] = role
    return roles


def normalize_role(role: str) -> str:
    value = role.strip().lower().replace("-", "_").replace(" ", "_")
    aliases = {
        "toc": "contents",
        "directory": "contents",
        "divider": "section",
        "screenshot": "evidence",
        "image": "evidence",
        "end": "closing",
    }
    value = aliases.get(value, value)
    if value not in ROLE_CHECKS:
        raise ValueError(f"Unknown slide role: {role!r}")
    return value


def infer_roles(
    indexes: list[int],
    explicit_roles: dict[int, str],
    dark_slides: set[int],
    evidence_slides: set[int],
) -> dict[int, str]:
    roles: dict[int, str] = {}
    last_index = max(indexes) if indexes else None
    for index in indexes:
        if index in explicit_roles:
            roles[index] = explicit_roles[index]
        elif index == 1:
            roles[index] = "cover"
        elif index == 2:
            roles[index] = "contents"
        elif last_index is not None and index == last_index:
            roles[index] = "closing"
        elif index in evidence_slides:
            roles[index] = "evidence"
        elif index in dark_slides:
            roles[index] = "section"
        else:
            roles[index] = "content"
    return roles


def load_json(path: Path | None) -> dict[str, Any] | None:
    if path is None:
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def qa_status_summary(report: dict[str, Any] | None) -> dict[str, Any] | None:
    if not report:
        return None
    return {
        "status": report.get("status"),
        "issues": report.get("summary", {}).get("issues", {}),
        "summary": {
            key: value
            for key, value in report.get("summary", {}).items()
            if key != "issues"
        },
    }


def source_captions_by_slide(pptx_qa: dict[str, Any] | None) -> dict[int, list[str]]:
    captions: dict[int, list[str]] = {}
    if not pptx_qa:
        return captions
    for slide in pptx_qa.get("slides", []):
        index = slide.get("index")
        if isinstance(index, int):
            captions[index] = list(slide.get("source_captions") or [])
    return captions


def preview_metrics_by_slide(preview_report: dict[str, Any] | None) -> dict[int, dict[str, Any]]:
    metrics: dict[int, dict[str, Any]] = {}
    if not preview_report:
        return metrics
    for image in preview_report.get("images", []):
        index = image.get("index")
        if isinstance(index, int):
            metrics[index] = {
                "white_ratio": image.get("white_ratio"),
                "blue_ratio": image.get("blue_ratio"),
                "non_white_ratio": image.get("non_white_ratio"),
                "luminance_std": image.get("luminance_std"),
                "content_bbox": image.get("content_bbox"),
            }
    return metrics


def build_packet(
    preview_dir: Path,
    *,
    output_dir: Path,
    pptx_qa: dict[str, Any] | None,
    preview_report: dict[str, Any] | None,
    dark_slides: set[int],
    evidence_slides: set[int],
    explicit_roles: dict[int, str],
    deck_title: str,
    reviewer_note: str,
    columns: int,
    thumb_width: int,
) -> dict[str, Any]:
    images = preview_qa.preview_images(preview_dir)
    indexes = [index for index, _path in images]
    roles = infer_roles(indexes, explicit_roles, dark_slides, evidence_slides)
    captions = source_captions_by_slide(pptx_qa)
    metrics = preview_metrics_by_slide(preview_report)

    slides = []
    for index, path in images:
        role = roles[index]
        slides.append({
            "index": index,
            "role": role,
            "image": str(path.resolve()),
            "source_captions": captions.get(index, []),
            "preview_metrics": metrics.get(index, {}),
            "checks": ROLE_CHECKS[role],
            "review_result": {
                "status": "unreviewed",
                "issues": [],
                "fix_notes": "",
            },
        })

    output_dir.mkdir(parents=True, exist_ok=True)
    contact_sheet = output_dir / "vision-contact-sheet.png"
    contact_sheet_created = create_contact_sheet(
        images,
        roles,
        contact_sheet,
        columns=columns,
        thumb_width=thumb_width,
    )

    packet = {
        "deck_title": deck_title,
        "preview_dir": str(preview_dir.resolve()),
        "output_dir": str(output_dir.resolve()),
        "reviewer_note": reviewer_note,
        "pptx_qa": qa_status_summary(pptx_qa),
        "preview_qa": qa_status_summary(preview_report),
        "contact_sheet": str(contact_sheet.resolve()) if contact_sheet_created else None,
        "slides": slides,
        "review_status": "unreviewed",
        "final_gate": [
            "PPTX QA status is pass.",
            "Preview-image QA status is pass.",
            "Every slide is visually reviewed with no blocking issues.",
            "Any issue found during review is fixed and the affected slides are re-exported.",
        ],
    }
    packet["packet_issues"] = packet_issues(packet)
    packet["packet_status"] = "pass" if not packet["packet_issues"] else "warning"
    return packet


def packet_issues(packet: dict[str, Any]) -> list[dict[str, Any]]:
    issues: list[dict[str, Any]] = []
    if not packet.get("slides"):
        issues.append({"severity": "error", "message": "No preview slides were found in the packet."})
    if not packet.get("contact_sheet"):
        issues.append({"severity": "warning", "message": "Contact sheet was not generated; PNG previews may be unsupported."})
    for label, key in [("PPTX QA", "pptx_qa"), ("Preview-image QA", "preview_qa")]:
        summary = packet.get(key)
        if summary and summary.get("status") != "pass":
            issues.append({
                "severity": "error",
                "message": f"{label} status is not pass.",
                "status": summary.get("status"),
            })
    return issues


def create_contact_sheet(
    images: list[tuple[int, Path]],
    roles: dict[int, str],
    output_path: Path,
    *,
    columns: int = 3,
    thumb_width: int = 320,
) -> bool:
    if not images:
        return False
    columns = max(1, columns)
    padding = 24
    gap = 18
    border = 8
    loaded = []
    for index, path in images:
        try:
            width, height, pixels = read_png_rgba(path)
        except Exception:
            return False
        thumb_height = max(1, round(thumb_width * height / width))
        loaded.append((index, width, height, pixels, thumb_width, thumb_height))

    rows = math.ceil(len(loaded) / columns)
    cell_width = thumb_width + border * 2
    cell_height = max(item[5] for item in loaded) + border * 2
    canvas_width = padding * 2 + columns * cell_width + (columns - 1) * gap
    canvas_height = padding * 2 + rows * cell_height + (rows - 1) * gap
    canvas = bytearray([246, 248, 252, 255] * canvas_width * canvas_height)

    for position, item in enumerate(loaded):
        index, width, height, pixels, target_width, target_height = item
        row = position // columns
        col = position % columns
        x = padding + col * (cell_width + gap)
        y = padding + row * (cell_height + gap)
        role = roles.get(index, "content")
        draw_rect(canvas, canvas_width, x, y, cell_width, cell_height, ROLE_COLORS.get(role, ROLE_COLORS["content"]))
        draw_rect(canvas, canvas_width, x + border, y + border, target_width, target_height, (255, 255, 255, 255))
        paste_scaled(canvas, canvas_width, x + border, y + border, target_width, target_height, width, height, pixels)
        draw_number_badge(canvas, canvas_width, x + border + 10, y + border + 10, index)

    write_png_rgba(output_path, canvas_width, canvas_height, canvas)
    return True


def read_png_rgba(path: Path) -> tuple[int, int, bytearray]:
    data = path.read_bytes()
    width, height, bit_depth, color_type, interlace, palette, idat = preview_qa.read_png_chunks(data)
    if bit_depth != 8 or interlace != 0 or color_type not in {0, 2, 3, 4, 6}:
        raise ValueError("Unsupported PNG format for contact sheet.")
    bpp = {0: 1, 2: 3, 3: 1, 4: 2, 6: 4}[color_type]
    row_bytes = width * bpp
    raw = zlib.decompress(idat)
    pixels = bytearray(width * height * 4)
    previous = bytearray(row_bytes)
    offset = 0
    for y in range(height):
        filter_type = raw[offset]
        offset += 1
        row = bytearray(raw[offset:offset + row_bytes])
        offset += row_bytes
        preview_qa.unfilter_row(row, previous, bpp, filter_type)
        for x in range(width):
            r, g, b, a = preview_qa.pixel_rgba(row, x * bpp, color_type, palette)
            dest = (y * width + x) * 4
            pixels[dest:dest + 4] = bytes((r, g, b, a))
        previous = row
    return width, height, pixels


def draw_rect(canvas: bytearray, canvas_width: int, x: int, y: int, width: int, height: int, color: tuple[int, int, int, int]) -> None:
    for yy in range(y, y + height):
        start = (yy * canvas_width + x) * 4
        end = start + width * 4
        canvas[start:end] = bytes(color) * width


def paste_scaled(
    canvas: bytearray,
    canvas_width: int,
    dest_x: int,
    dest_y: int,
    dest_w: int,
    dest_h: int,
    src_w: int,
    src_h: int,
    src_pixels: bytearray,
) -> None:
    for y in range(dest_h):
        sy = min(src_h - 1, int(y * src_h / dest_h))
        for x in range(dest_w):
            sx = min(src_w - 1, int(x * src_w / dest_w))
            src = (sy * src_w + sx) * 4
            dest = ((dest_y + y) * canvas_width + dest_x + x) * 4
            alpha = src_pixels[src + 3] / 255
            inv = 1 - alpha
            canvas[dest] = round(src_pixels[src] * alpha + canvas[dest] * inv)
            canvas[dest + 1] = round(src_pixels[src + 1] * alpha + canvas[dest + 1] * inv)
            canvas[dest + 2] = round(src_pixels[src + 2] * alpha + canvas[dest + 2] * inv)
            canvas[dest + 3] = 255


def draw_number_badge(canvas: bytearray, canvas_width: int, x: int, y: int, index: int) -> None:
    draw_rect(canvas, canvas_width, x, y, 46, 24, (255, 255, 255, 235))
    draw_text_digits(canvas, canvas_width, x + 8, y + 5, f"{index:02d}", scale=3, color=(31, 55, 107, 255))


def draw_text_digits(
    canvas: bytearray,
    canvas_width: int,
    x: int,
    y: int,
    text: str,
    *,
    scale: int,
    color: tuple[int, int, int, int],
) -> None:
    cursor = x
    for char in text:
        pattern = DIGIT_FONT.get(char)
        if pattern:
            for py, row in enumerate(pattern):
                for px, value in enumerate(row):
                    if value == "1":
                        draw_rect(canvas, canvas_width, cursor + px * scale, y + py * scale, scale, scale, color)
            cursor += (len(pattern[0]) + 1) * scale
        else:
            cursor += 2 * scale


def write_png_rgba(path: Path, width: int, height: int, pixels: bytearray) -> None:
    import binascii

    rows = []
    stride = width * 4
    for y in range(height):
        rows.append(b"\x00" + bytes(pixels[y * stride:(y + 1) * stride]))
    raw = zlib.compress(b"".join(rows))

    def chunk(name: bytes, payload: bytes) -> bytes:
        crc = binascii.crc32(name + payload) & 0xFFFFFFFF
        return len(payload).to_bytes(4, "big") + name + payload + crc.to_bytes(4, "big")

    ihdr = width.to_bytes(4, "big") + height.to_bytes(4, "big") + bytes([8, 6, 0, 0, 0])
    path.write_bytes(
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", ihdr)
        + chunk(b"IDAT", raw)
        + chunk(b"IEND", b"")
    )


def write_markdown(packet: dict[str, Any], output_path: Path) -> None:
    lines = [
        f"# Vision Review Packet: {packet['deck_title']}",
        "",
        "Use this packet for the final visual review after PPTX QA and preview-image QA pass.",
        "Assume issues exist; the reviewer should actively look for them, not merely confirm the deck.",
        "",
        "## QA Summary",
        "",
    ]
    for label, key in [("PPTX QA", "pptx_qa"), ("Preview-image QA", "preview_qa")]:
        summary = packet.get(key)
        if summary:
            lines.append(f"- {label}: `{summary.get('status')}`; issues: `{summary.get('issues')}`")
        else:
            lines.append(f"- {label}: not provided")
    if packet.get("contact_sheet"):
        lines.extend(["", "## Contact Sheet", "", f"![Vision contact sheet]({Path(packet['contact_sheet']).as_posix()})"])
    lines.extend([
        "",
        "## Global Checks",
        "",
        "- Role separation is correct: cover, contents, section, content/evidence, closing.",
        "- No title/body text collision, clipping, or unreadable wrapping.",
        "- No unexplained white background blocks on dark template pages.",
        "- Evidence pages have legible screenshots, source captions, and blue highlights that point to the cited region.",
        "- The deck feels like one SCUT-blue system rather than mixed masters.",
        "",
        "## Slide Checks",
        "",
    ])
    for slide in packet["slides"]:
        image = Path(slide["image"]).as_posix()
        lines.extend([
            f"### Slide {slide['index']:02d} - {slide['role']}",
            "",
            f"![Slide {slide['index']:02d}]({image})",
            "",
            "Checks:",
        ])
        for check in slide["checks"]:
            lines.append(f"- [ ] {check}")
        if slide.get("source_captions"):
            lines.append(f"- Source captions: `{'; '.join(slide['source_captions'])}`")
        metrics = slide.get("preview_metrics") or {}
        if metrics:
            compact = {key: metrics.get(key) for key in ("white_ratio", "blue_ratio", "non_white_ratio", "luminance_std")}
            lines.append(f"- Preview metrics: `{compact}`")
        lines.extend(["- Review finding:", "- Required fix:", ""])
    output_path.write_text("\n".join(lines), encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build a structured Vision/human review packet for SCUT paper PPT previews.")
    parser.add_argument("preview_dir", type=Path, help="Directory containing exported slide PNG/JPEG previews.")
    parser.add_argument("--output-dir", type=Path, required=True, help="Directory to write the review packet.")
    parser.add_argument("--pptx-qa", type=Path, help="PPTX bottom-layer QA JSON report.")
    parser.add_argument("--preview-qa", type=Path, help="Preview-image QA JSON report.")
    parser.add_argument("--dark-slide", action="append", type=int, default=[], help="Slide index expected to use a dark/template background.")
    parser.add_argument("--evidence-slide", action="append", type=int, default=[], help="Slide index expected to be an evidence/screenshot page.")
    parser.add_argument("--slide-role", action="append", default=[], help="Explicit slide role as INDEX:cover|contents|section|content|evidence|closing.")
    parser.add_argument("--deck-title", default="SCUT paper presentation", help="Human-readable deck title for the packet.")
    parser.add_argument("--reviewer-note", default="", help="Optional note to include in the JSON packet.")
    parser.add_argument("--columns", type=int, default=3, help="Contact sheet columns.")
    parser.add_argument("--thumb-width", type=int, default=320, help="Contact sheet thumbnail width in pixels.")
    parser.add_argument("--strict", action="store_true", help="Exit non-zero unless packet status is pass.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    packet = build_packet(
        args.preview_dir,
        output_dir=args.output_dir,
        pptx_qa=load_json(args.pptx_qa),
        preview_report=load_json(args.preview_qa),
        dark_slides=set(args.dark_slide),
        evidence_slides=set(args.evidence_slide),
        explicit_roles=parse_slide_roles(args.slide_role),
        deck_title=args.deck_title,
        reviewer_note=args.reviewer_note,
        columns=args.columns,
        thumb_width=args.thumb_width,
    )
    json_path = args.output_dir / "vision-review-packet.json"
    markdown_path = args.output_dir / "vision-review-prompt.md"
    json_path.write_text(json.dumps(packet, ensure_ascii=False, indent=2), encoding="utf-8")
    write_markdown(packet, markdown_path)
    print(json.dumps({
        "status": packet.get("packet_status", "warning"),
        "issues": packet.get("packet_issues", []),
        "packet": str(json_path),
        "prompt": str(markdown_path),
        "contact_sheet": packet.get("contact_sheet"),
        "slide_count": len(packet.get("slides", [])),
    }, ensure_ascii=False, indent=2))
    return 1 if args.strict and packet.get("packet_status") != "pass" else 0


if __name__ == "__main__":
    raise SystemExit(main())
