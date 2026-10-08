#!/usr/bin/env python
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import zlib
from pathlib import Path
from typing import Any


PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg"}


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
    for issue in report.get("issues", []):
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


def natural_key(text: str) -> tuple[Any, ...]:
    return tuple(int(chunk) if chunk.isdigit() else chunk for chunk in re.split(r"(\d+)", text))


def slide_index(filename: str) -> int | None:
    stem = Path(filename).stem
    match = re.search(r"(?i)(?:^|[-_\s])slide[-_\s]?0*(\d+)(?:$|[-_\s])", stem)
    if not match:
        match = re.search(r"(?i)^slide[-_\s]?0*(\d+)$", stem)
    if not match:
        match = re.search(r"0*(\d+)$", stem)
    if not match:
        return None
    try:
        return int(match.group(1))
    except ValueError:
        return None


def preview_images(preview_dir: Path) -> list[tuple[int, Path]]:
    images: list[tuple[int, Path]] = []
    fallback = 1
    for child in preview_dir.iterdir():
        if not child.is_file() or child.suffix.lower() not in IMAGE_EXTENSIONS:
            continue
        index = slide_index(child.name)
        if index is None:
            index = fallback
            fallback += 1
        images.append((index, child))
    return sorted(images, key=lambda item: (item[0], natural_key(item[1].name)))


def qa_preview(
    preview_dir: Path,
    *,
    expected_slides: int | None = None,
    dark_slides: list[int] | None = None,
    evidence_slides: list[int] | None = None,
    min_width: int = 800,
    min_height: int = 450,
    max_dark_white_ratio: float = 0.38,
    min_evidence_blue_ratio: float = 0.004,
) -> dict[str, Any]:
    issues: list[dict[str, Any]] = []
    dark_set = {int(index) for index in dark_slides or []}
    evidence_set = {int(index) for index in evidence_slides or []}
    report: dict[str, Any] = {
        "path": str(preview_dir.resolve()),
        "exists": preview_dir.is_dir(),
        "status": "error",
        "summary": {},
        "images": [],
        "issues": issues,
    }
    if not preview_dir.is_dir():
        add_issue(issues, "error", "preview", "Preview directory does not exist.", part=str(preview_dir))
        return finalize(report)

    images = preview_images(preview_dir)
    if expected_slides is not None and len(images) != expected_slides:
        add_issue(
            issues,
            "error",
            "preview",
            "Preview image count does not match expected slide count.",
            details={"expected": expected_slides, "actual": len(images)},
        )

    found = {index for index, _path in images}
    for index in sorted((dark_set | evidence_set) - found):
        add_issue(issues, "error", "preview", "Expected preview image for a role-specific slide was not found.", slide=index)

    base_size: tuple[int, int] | None = None
    for index, path in images:
        image = inspect_image(path, index, issues)
        report["images"].append(image)
        width = image.get("width_px")
        height = image.get("height_px")
        if not isinstance(width, int) or not isinstance(height, int):
            continue
        if width < min_width or height < min_height:
            add_issue(
                issues,
                "warning",
                "preview",
                "Preview image resolution is lower than the configured minimum.",
                slide=index,
                part=str(path),
                details={"width_px": width, "height_px": height, "min_width_px": min_width, "min_height_px": min_height},
            )
        aspect_ratio = width / height if height else 0
        if abs(aspect_ratio - (16 / 9)) > 0.04:
            add_issue(
                issues,
                "warning",
                "preview",
                "Preview image aspect ratio is not close to 16:9.",
                slide=index,
                part=str(path),
                details={"aspect_ratio": round(aspect_ratio, 4)},
            )
        if base_size is None:
            base_size = (width, height)
        elif base_size != (width, height):
            add_issue(
                issues,
                "warning",
                "preview",
                "Preview images do not all share the same dimensions.",
                slide=index,
                part=str(path),
                details={"expected": base_size, "actual": (width, height)},
            )

        non_white_ratio = image.get("non_white_ratio")
        luma_std = image.get("luminance_std")
        if isinstance(non_white_ratio, float) and isinstance(luma_std, float):
            if non_white_ratio < 0.001 and luma_std < 2.0:
                add_issue(
                    issues,
                    "error",
                    "visual",
                    "Preview image appears blank or nearly blank.",
                    slide=index,
                    part=str(path),
                    details={"non_white_ratio": non_white_ratio, "luminance_std": luma_std},
                )

        if index in dark_set:
            white_ratio = image.get("white_ratio")
            if isinstance(white_ratio, float) and white_ratio > max_dark_white_ratio:
                add_issue(
                    issues,
                    "error",
                    "visual",
                    "Dark/template slide has too much white area; possible wrong master or white overlay.",
                    slide=index,
                    part=str(path),
                    details={"white_ratio": white_ratio, "max_allowed": max_dark_white_ratio},
                )
        if index in evidence_set:
            blue_ratio = image.get("blue_ratio")
            if isinstance(blue_ratio, float) and blue_ratio < min_evidence_blue_ratio:
                add_issue(
                    issues,
                    "error",
                    "visual",
                    "Evidence slide has little detectable SCUT-blue accent/highlight.",
                    slide=index,
                    part=str(path),
                    details={"blue_ratio": blue_ratio, "min_required": min_evidence_blue_ratio},
                )

    report["summary"] = {
        "image_count": len(images),
        "dark_slide_indexes": sorted(dark_set),
        "evidence_slide_indexes": sorted(evidence_set),
        "base_width_px": base_size[0] if base_size else None,
        "base_height_px": base_size[1] if base_size else None,
    }
    return finalize(report)


def inspect_image(path: Path, index: int, issues: list[dict[str, Any]]) -> dict[str, Any]:
    report: dict[str, Any] = {
        "index": index,
        "path": str(path.resolve()),
        "width_px": None,
        "height_px": None,
        "format": None,
        "pixel_metrics_available": False,
    }
    try:
        data = path.read_bytes()
        metrics = image_metrics(data)
    except Exception as exc:
        add_issue(issues, "warning", "preview", str(exc), slide=index, part=str(path))
        return report
    report.update(metrics)
    if not report.get("pixel_metrics_available"):
        add_issue(
            issues,
            "warning",
            "preview",
            "Pixel-level preview QA is unavailable for this image format or encoding.",
            slide=index,
            part=str(path),
            details={"format": report.get("format")},
        )
    return report


def image_metrics(data: bytes) -> dict[str, Any]:
    if data.startswith(PNG_SIGNATURE):
        return png_metrics(data)
    if data.startswith(b"\xff\xd8"):
        width, height = jpeg_dimensions(data)
        return {"format": "jpeg", "width_px": width, "height_px": height, "pixel_metrics_available": False}
    raise ValueError("Preview image is not a supported PNG/JPEG file.")


def png_metrics(data: bytes) -> dict[str, Any]:
    width, height, bit_depth, color_type, interlace, palette, idat = read_png_chunks(data)
    base: dict[str, Any] = {
        "format": "png",
        "width_px": width,
        "height_px": height,
        "bit_depth": bit_depth,
        "color_type": color_type,
        "interlace": interlace,
        "pixel_metrics_available": False,
    }
    if bit_depth != 8 or interlace != 0 or color_type not in {0, 2, 3, 4, 6}:
        return base
    bpp = {0: 1, 2: 3, 3: 1, 4: 2, 6: 4}[color_type]
    row_bytes = width * bpp
    raw = zlib.decompress(idat)
    expected = (row_bytes + 1) * height
    if len(raw) < expected:
        raise ValueError("PNG preview image data is truncated.")

    previous = bytearray(row_bytes)
    offset = 0
    total = width * height
    white = non_white = blue = dark = 0
    luma_sum = 0.0
    luma_sq_sum = 0.0
    sampled = set()
    sample_step = max(1, total // 5000)
    pixel_index = 0
    bbox = [width, height, -1, -1]

    for y in range(height):
        filter_type = raw[offset]
        offset += 1
        row = bytearray(raw[offset:offset + row_bytes])
        offset += row_bytes
        unfilter_row(row, previous, bpp, filter_type)
        for x in range(width):
            start = x * bpp
            r, g, b, alpha = pixel_rgba(row, start, color_type, palette)
            if alpha < 8:
                r = g = b = 255
            luma = 0.2126 * r + 0.7152 * g + 0.0722 * b
            luma_sum += luma
            luma_sq_sum += luma * luma
            if r >= 245 and g >= 245 and b >= 245:
                white += 1
            else:
                non_white += 1
                bbox[0] = min(bbox[0], x)
                bbox[1] = min(bbox[1], y)
                bbox[2] = max(bbox[2], x)
                bbox[3] = max(bbox[3], y)
            if max(r, g, b) <= 45:
                dark += 1
            if is_scut_blue(r, g, b):
                blue += 1
            if pixel_index % sample_step == 0:
                sampled.add((r, g, b))
            pixel_index += 1
        previous = row

    mean = luma_sum / total if total else 0.0
    variance = max(0.0, (luma_sq_sum / total if total else 0.0) - mean * mean)
    base.update({
        "pixel_metrics_available": True,
        "white_ratio": round(white / total, 6) if total else 0.0,
        "non_white_ratio": round(non_white / total, 6) if total else 0.0,
        "blue_ratio": round(blue / total, 6) if total else 0.0,
        "dark_ratio": round(dark / total, 6) if total else 0.0,
        "luminance_mean": round(mean, 3),
        "luminance_std": round(variance ** 0.5, 3),
        "sampled_unique_colors": len(sampled),
        "content_bbox": None if bbox[2] < 0 else {"x": bbox[0], "y": bbox[1], "width": bbox[2] - bbox[0] + 1, "height": bbox[3] - bbox[1] + 1},
    })
    return base


def read_png_chunks(data: bytes) -> tuple[int, int, int, int, int, list[tuple[int, int, int]], bytes]:
    offset = len(PNG_SIGNATURE)
    width = height = bit_depth = color_type = interlace = 0
    palette: list[tuple[int, int, int]] = []
    idat: list[bytes] = []
    while offset + 8 <= len(data):
        length = int.from_bytes(data[offset:offset + 4], "big")
        chunk_type = data[offset + 4:offset + 8]
        offset += 8
        payload = data[offset:offset + length]
        offset += length + 4
        if chunk_type == b"IHDR":
            width = int.from_bytes(payload[0:4], "big")
            height = int.from_bytes(payload[4:8], "big")
            bit_depth = payload[8]
            color_type = payload[9]
            interlace = payload[12]
        elif chunk_type == b"PLTE":
            palette = [(payload[i], payload[i + 1], payload[i + 2]) for i in range(0, len(payload) - 2, 3)]
        elif chunk_type == b"IDAT":
            idat.append(payload)
        elif chunk_type == b"IEND":
            break
    if not width or not height or not idat:
        raise ValueError("PNG preview image is missing IHDR or IDAT data.")
    return width, height, bit_depth, color_type, interlace, palette, b"".join(idat)


def unfilter_row(row: bytearray, previous: bytearray, bpp: int, filter_type: int) -> None:
    for i, value in enumerate(row):
        left = row[i - bpp] if i >= bpp else 0
        up = previous[i] if previous else 0
        up_left = previous[i - bpp] if previous and i >= bpp else 0
        if filter_type == 0:
            recon = value
        elif filter_type == 1:
            recon = value + left
        elif filter_type == 2:
            recon = value + up
        elif filter_type == 3:
            recon = value + ((left + up) // 2)
        elif filter_type == 4:
            recon = value + paeth(left, up, up_left)
        else:
            raise ValueError(f"Unsupported PNG filter type: {filter_type}")
        row[i] = recon & 0xFF


def pixel_rgba(row: bytearray, start: int, color_type: int, palette: list[tuple[int, int, int]]) -> tuple[int, int, int, int]:
    if color_type == 0:
        value = row[start]
        return value, value, value, 255
    if color_type == 2:
        return row[start], row[start + 1], row[start + 2], 255
    if color_type == 3:
        index = row[start]
        if index < len(palette):
            r, g, b = palette[index]
            return r, g, b, 255
        return 0, 0, 0, 255
    if color_type == 4:
        value = row[start]
        return value, value, value, row[start + 1]
    return row[start], row[start + 1], row[start + 2], row[start + 3]


def paeth(left: int, up: int, up_left: int) -> int:
    p = left + up - up_left
    pa = abs(p - left)
    pb = abs(p - up)
    pc = abs(p - up_left)
    if pa <= pb and pa <= pc:
        return left
    if pb <= pc:
        return up
    return up_left


def is_scut_blue(r: int, g: int, b: int) -> bool:
    return b >= 110 and (b - r) >= 35 and (b - g) >= 20 and g >= 45


def jpeg_dimensions(data: bytes) -> tuple[int | None, int | None]:
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


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run deterministic QA against exported slide preview images.")
    parser.add_argument("preview_dir", type=Path, help="Directory containing exported slide PNG/JPEG previews.")
    parser.add_argument("--expected-slides", type=int, help="Expected preview image count.")
    parser.add_argument("--dark-slide", action="append", type=int, default=[], help="Slide index expected to use a dark/template background.")
    parser.add_argument("--evidence-slide", action="append", type=int, default=[], help="Slide index expected to contain SCUT-blue evidence highlighting.")
    parser.add_argument("--min-width", type=int, default=800, help="Minimum preview width in pixels.")
    parser.add_argument("--min-height", type=int, default=450, help="Minimum preview height in pixels.")
    parser.add_argument("--max-dark-white-ratio", type=float, default=0.38, help="Maximum white-pixel ratio allowed on dark slides.")
    parser.add_argument("--min-evidence-blue-ratio", type=float, default=0.004, help="Minimum blue-pixel ratio required on evidence slides.")
    parser.add_argument("--report", type=Path, help="Write the full JSON report.")
    parser.add_argument("--strict", action="store_true", help="Exit non-zero unless status is pass.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    report = qa_preview(
        args.preview_dir,
        expected_slides=args.expected_slides,
        dark_slides=args.dark_slide,
        evidence_slides=args.evidence_slide,
        min_width=args.min_width,
        min_height=args.min_height,
        max_dark_white_ratio=args.max_dark_white_ratio,
        min_evidence_blue_ratio=args.min_evidence_blue_ratio,
    )
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
