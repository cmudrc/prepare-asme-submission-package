#!/usr/bin/env python3
"""Build and audit an ASME journal final-submission package.

The program deliberately refuses to manufacture image resolution or to infer
ambiguous multipart layouts.  It creates QA artifacts even when a hard check
fails, but creates the upload ZIP only after every hard check passes.
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import re
import shutil
import subprocess
import sys
import tempfile
import uuid
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable

try:
    from lxml import etree
    from PIL import Image
    from pypdf import PdfReader
except ImportError as exc:  # pragma: no cover - exercised only outside Codex runtime
    raise SystemExit(
        "Missing dependency. Install lxml, Pillow, and pypdf, then run again: "
        f"{exc}"
    )


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"w": W, "r": R, "a": A, "wp": WP, "pr": PKG_REL}
RASTER_EXTENSIONS = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"}
VECTOR_EXTENSIONS = {".eps", ".pdf"}
GRAPHIC_EXTENSIONS = RASTER_EXTENSIONS | VECTOR_EXTENSIONS | {
    ".svg",
    ".emf",
    ".wmf",
}
DEFAULT_THRESHOLDS = {"photo": 266.0, "composite": 500.0, "linework": 900.0}
MAX_GRAPHIC_BYTES = 15 * 1024 * 1024
FIGURE_CAPTION_RE = re.compile(
    r"^\s*fig(?:ure)?s?\.?\s*(\d+)([a-z])?\s*(?:[.:\-]|$)", re.IGNORECASE
)
FIGURE_CALLOUT_RE = re.compile(
    r"\bfig(?:ure)?s?\.?\s*(\d+)([a-z])?\b", re.IGNORECASE
)
FIGURE_GROUP_CALLOUT_RE = re.compile(
    r"\bfig(?:ure)?s?\.?\s*(\d+[a-z]?(?:\s*(?:,|and|&|to|[-–—])\s*\d+[a-z]?)+)",
    re.IGNORECASE,
)


@dataclass
class Finding:
    level: str
    code: str
    message: str


@dataclass
class Context:
    args: argparse.Namespace
    root: Path
    upload: Path
    qa: Path
    findings: list[Finding] = field(default_factory=list)
    figures: list[dict] = field(default_factory=list)
    source_text: str = ""
    complete_pdf: Path | None = None
    text_only_render: Path | None = None
    latex_main_relative: Path | None = None

    def error(self, code: str, message: str) -> None:
        self.findings.append(Finding("ERROR", code, message))

    def warn(self, code: str, message: str) -> None:
        self.findings.append(Finding("WARNING", code, message))

    def note(self, code: str, message: str) -> None:
        self.findings.append(Finding("NOTE", code, message))

    @property
    def blocked(self) -> bool:
        return any(f.level == "ERROR" for f in self.findings)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build a validated ASME journal final-submission package."
    )
    parser.add_argument("source", type=Path, help="Authoritative .docx or Overleaf .zip")
    parser.add_argument("--paper-id", required=True)
    parser.add_argument(
        "--journal",
        default="asme",
        help="ASME journal name or short code (for example, jmd, jcise, or jvse)",
    )
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--main", help="Root .tex path within an Overleaf ZIP")
    parser.add_argument("--full-pdf", type=Path, help="Accepted complete PDF for LaTeX input")
    parser.add_argument("--due-date", help="Editorial due date in YYYY-MM-DD form")
    parser.add_argument(
        "--multipart",
        choices=("separate", "combined"),
        default="separate",
        help="Controlling editorial instruction for multipart figures",
    )
    parser.add_argument(
        "--figure-kind",
        action="append",
        default=[],
        metavar="FIG=KIND",
        help="Classify a raster as photo, composite, or linework",
    )
    parser.add_argument(
        "--figure-width",
        action="append",
        default=[],
        metavar="FIG=INCHES",
        help="Printed width for a LaTeX figure whose width cannot be inferred",
    )
    parser.add_argument(
        "--text-width",
        type=float,
        help="LaTeX full text-block width in inches when it cannot be inferred from the class",
    )
    parser.add_argument(
        "--column-width",
        type=float,
        help="LaTeX single-column width in inches when it cannot be inferred from the class",
    )
    parser.add_argument(
        "--required-dpi",
        type=float,
        help="Paper-specific minimum that overrides type-based ASME thresholds",
    )
    parser.add_argument(
        "--require-section",
        action="append",
        default=[],
        help="Text or heading that must appear in the authoritative source",
    )
    parser.add_argument(
        "--instruction",
        action="append",
        default=[],
        help="Controlling acceptance-letter or editorial instruction to record",
    )
    parser.add_argument(
        "--portal-rule",
        action="append",
        default=[],
        help="Portal filename, upload, or permission rule to record",
    )
    parser.add_argument(
        "--supplemental-file",
        action="append",
        default=[],
        type=Path,
        help="Additional production file to include under upload/supplemental",
    )
    return parser.parse_args(argv)


def parse_assignments(values: Iterable[str], allowed: set[str] | None = None) -> dict[str, str]:
    result: dict[str, str] = {}
    for value in values:
        if "=" not in value:
            raise ValueError(f"Expected FIG=VALUE, got {value!r}")
        key, raw = (part.strip() for part in value.split("=", 1))
        key = normalize_figure_id(key)
        if not key:
            raise ValueError(f"Missing figure identifier in {value!r}")
        if allowed is not None and raw not in allowed:
            raise ValueError(f"Expected one of {sorted(allowed)}, got {raw!r}")
        result[key] = raw
    return result


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def normalize_figure_id(value: str) -> str:
    match = re.search(r"(\d+)\s*([a-z])?", value.strip(), re.IGNORECASE)
    return f"{int(match.group(1))}{(match.group(2) or '').lower()}" if match else ""


def figure_sort_key(value: str) -> tuple[int, str]:
    match = re.fullmatch(r"(\d+)([a-z]?)", value)
    return (int(match.group(1)), match.group(2)) if match else (10**9, value)


def safe_name(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", value).strip("._")
    return cleaned or "submission"


def safe_extract_zip(source: Path, target: Path, ctx: Context) -> bool:
    try:
        with zipfile.ZipFile(source) as archive:
            root = target.resolve()
            for info in archive.infolist():
                destination = (target / info.filename).resolve()
                if destination != root and root not in destination.parents:
                    ctx.error("UNSAFE_ZIP", f"ZIP member escapes the project directory: {info.filename}")
                    return False
            archive.extractall(target)
        return True
    except (zipfile.BadZipFile, OSError) as exc:
        ctx.error("INVALID_ZIP", f"Could not extract {source.name}: {exc}")
        return False


def xml_parse(data: bytes) -> etree._Element:
    return etree.fromstring(data, parser=etree.XMLParser(remove_blank_text=False, recover=False))


def remove_elements(root: etree._Element, xpath: str, namespaces: dict[str, str] = NS) -> int:
    count = 0
    for element in list(root.xpath(xpath, namespaces=namespaces)):
        parent = element.getparent()
        if parent is not None:
            parent.remove(element)
            count += 1
    return count


def unwrap_elements(root: etree._Element, xpath: str) -> int:
    count = 0
    for element in list(root.xpath(xpath, namespaces=NS)):
        parent = element.getparent()
        if parent is None:
            continue
        index = parent.index(element)
        for child in list(element):
            element.remove(child)
            parent.insert(index, child)
            index += 1
        parent.remove(element)
        count += 1
    return count


def sanitize_word_xml(data: bytes, remove_graphics: bool) -> tuple[bytes, dict[str, int]]:
    root = xml_parse(data)
    counts = {
        "insertions": 0,
        "deletions": 0,
        "deleted_paragraphs": 0,
        "format_changes": 0,
        "comments": 0,
        "graphics": 0,
    }
    deletion_paragraphs = set(
        root.xpath(".//w:del/ancestor::w:p | .//w:moveFrom/ancestor::w:p", namespaces=NS)
    )
    counts["deletions"] += remove_elements(root, ".//w:del | .//w:moveFrom")
    meaningful = (
        ".//w:t[normalize-space()] | .//w:tab | .//w:br | .//w:drawing | .//w:pict | "
        ".//w:object | .//w:fldChar | .//w:instrText | .//w:sectPr"
    )
    for paragraph in deletion_paragraphs:
        parent = paragraph.getparent()
        if parent is not None and not paragraph.xpath(meaningful, namespaces=NS):
            parent.remove(paragraph)
            counts["deleted_paragraphs"] += 1
    counts["insertions"] += unwrap_elements(root, ".//w:ins | .//w:moveTo")
    counts["format_changes"] += remove_elements(
        root,
        ".//w:pPrChange | .//w:rPrChange | .//w:tblPrChange | .//w:trPrChange | "
        ".//w:tcPrChange | .//w:sectPrChange | .//w:tblGridChange | .//w:numberingChange",
    )
    remove_elements(
        root,
        ".//w:moveFromRangeStart | .//w:moveFromRangeEnd | "
        ".//w:moveToRangeStart | .//w:moveToRangeEnd",
    )
    counts["comments"] += remove_elements(
        root, ".//w:commentRangeStart | .//w:commentRangeEnd | .//w:commentReference"
    )
    remove_elements(root, ".//w:trackRevisions")
    if remove_graphics:
        counts["graphics"] += remove_elements(root, ".//w:drawing | .//w:pict | .//w:object")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True), counts


def sanitize_relationships(data: bytes, remove_graphics: bool) -> bytes:
    root = xml_parse(data)
    for rel in list(root):
        rel_type = (rel.get("Type") or "").lower()
        target = (rel.get("Target") or "").lower()
        if rel_type.endswith("/comments") or "comment" in rel_type:
            root.remove(rel)
        elif remove_graphics and (rel_type.endswith("/image") or "media/" in target):
            root.remove(rel)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def sanitize_content_types(data: bytes, remove_graphics: bool) -> bytes:
    root = xml_parse(data)
    for element in list(root):
        part = (element.get("PartName") or "").lower()
        extension = (element.get("Extension") or "").lower()
        if "comment" in part or (remove_graphics and extension in {e.lstrip(".") for e in GRAPHIC_EXTENSIONS}):
            root.remove(element)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def sanitize_docx(source: Path, destination: Path, remove_graphics: bool) -> dict[str, int]:
    totals = {
        "insertions": 0,
        "deletions": 0,
        "deleted_paragraphs": 0,
        "format_changes": 0,
        "comments": 0,
        "graphics": 0,
    }
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            lower = item.filename.lower()
            if "comment" in lower or lower in {"word/people.xml", "word/person.xml"}:
                continue
            if remove_graphics and lower.startswith("word/media/"):
                continue
            data = zin.read(item.filename)
            if lower == "[content_types].xml":
                data = sanitize_content_types(data, remove_graphics)
            elif lower.endswith(".rels"):
                data = sanitize_relationships(data, remove_graphics)
            elif lower.startswith("word/") and lower.endswith(".xml"):
                try:
                    data, counts = sanitize_word_xml(data, remove_graphics)
                    for key, value in counts.items():
                        totals[key] += value
                except etree.XMLSyntaxError:
                    pass
            zout.writestr(item, data)
    return totals


def docx_visible_text(path: Path) -> str:
    text_parts: list[str] = []
    with zipfile.ZipFile(path) as archive:
        for name in sorted(
            n for n in archive.namelist() if n.startswith("word/") and n.endswith(".xml")
        ):
            try:
                root = xml_parse(archive.read(name))
            except etree.XMLSyntaxError:
                continue
            paragraphs = root.xpath(".//w:p", namespaces=NS)
            if paragraphs:
                text_parts.extend(paragraph_text(paragraph) for paragraph in paragraphs)
            else:
                text_parts.append("".join(root.xpath(".//w:t/text()", namespaces=NS)))
    return "\n".join(text_parts)


def docx_has_graphics(path: Path) -> bool:
    with zipfile.ZipFile(path) as archive:
        if any(name.lower().startswith("word/media/") for name in archive.namelist()):
            return True
        for name in archive.namelist():
            if name.startswith("word/") and name.endswith(".xml"):
                try:
                    root = xml_parse(archive.read(name))
                except etree.XMLSyntaxError:
                    continue
                if root.xpath(".//w:drawing | .//w:pict | .//w:object", namespaces=NS):
                    return True
    return False


def paragraph_text(paragraph: etree._Element) -> str:
    return "".join(paragraph.xpath(".//w:t/text()", namespaces=NS)).strip()


def image_width_inches(blip: etree._Element) -> float | None:
    node = blip
    while node is not None and node.tag not in {f"{{{WP}}}inline", f"{{{WP}}}anchor"}:
        node = node.getparent()
    if node is None:
        return None
    extents = node.xpath(".//wp:extent", namespaces=NS)
    if not extents:
        return None
    try:
        return float(extents[0].get("cx")) / 914400.0
    except (TypeError, ValueError):
        return None


def docx_relationships(archive: zipfile.ZipFile) -> dict[str, str]:
    root = xml_parse(archive.read("word/_rels/document.xml.rels"))
    return {
        rel.get("Id"): rel.get("Target")
        for rel in root
        if rel.get("Id") and rel.get("Target")
    }


def next_caption(paragraphs: list[etree._Element], index: int) -> tuple[str, str] | None:
    for candidate in paragraphs[index : index + 5]:
        text = paragraph_text(candidate)
        match = FIGURE_CAPTION_RE.match(text)
        if match:
            return normalize_figure_id("".join(match.groups(default=""))), text
        if text:
            break
    return None


def numbered_figure_paragraphs(root: etree._Element) -> list[etree._Element]:
    paragraphs = root.xpath(".//w:body//w:p", namespaces=NS)
    return [
        paragraph
        for index, paragraph in enumerate(paragraphs)
        if paragraph.xpath(".//a:blip[@r:embed]", namespaces=NS)
        and next_caption(paragraphs, index + 1)
    ]


def remove_numbered_docx_figures(source: Path, destination: Path) -> dict[str, int]:
    """Remove numbered production figures while retaining inline semantic graphics."""
    with zipfile.ZipFile(source) as archive:
        infos = archive.infolist()
        contents = {info.filename: archive.read(info.filename) for info in infos}

    document = xml_parse(contents["word/document.xml"])
    removed_relationship_ids: set[str] = set()
    removed_nodes = 0
    for paragraph in numbered_figure_paragraphs(document):
        for graphic in list(paragraph.xpath(".//w:drawing | .//w:pict | .//w:object", namespaces=NS)):
            removed_relationship_ids.update(
                graphic.xpath(".//a:blip/@r:embed", namespaces=NS)
            )
            parent = graphic.getparent()
            if parent is not None:
                parent.remove(graphic)
                removed_nodes += 1
    contents["word/document.xml"] = etree.tostring(
        document, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    removed_members: set[str] = set()
    rel_name = "word/_rels/document.xml.rels"
    relationships_removed = 0
    if rel_name in contents and removed_relationship_ids:
        rel_root = xml_parse(contents[rel_name])
        remaining_targets = {
            rel.get("Target")
            for rel in rel_root
            if rel.get("Id") not in removed_relationship_ids and rel.get("Target")
        }
        for rel in list(rel_root):
            if rel.get("Id") not in removed_relationship_ids:
                continue
            target = rel.get("Target")
            rel_root.remove(rel)
            relationships_removed += 1
            if target and target not in remaining_targets:
                removed_members.add((Path("word") / target).as_posix())
        contents[rel_name] = etree.tostring(
            rel_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as archive:
        for info in infos:
            if info.filename in removed_members:
                continue
            archive.writestr(info, contents[info.filename])
    return {
        "graphics": removed_nodes,
        "relationships": relationships_removed,
        "media": len(removed_members),
    }


def docx_has_numbered_figures(path: Path) -> bool:
    with zipfile.ZipFile(path) as archive:
        root = xml_parse(archive.read("word/document.xml"))
    return bool(numbered_figure_paragraphs(root))


def extract_docx_graphics(path: Path, ctx: Context) -> list[dict]:
    records: list[dict] = []
    with zipfile.ZipFile(path) as archive:
        root = xml_parse(archive.read("word/document.xml"))
        rels = docx_relationships(archive)
        paragraphs = root.xpath(".//w:body//w:p", namespaces=NS)
        for index, paragraph in enumerate(paragraphs):
            blips = paragraph.xpath(".//a:blip[@r:embed]", namespaces=NS)
            if not blips:
                continue
            caption = next_caption(paragraphs, index + 1)
            for blip in blips:
                rel_id = blip.get(f"{{{R}}}embed")
                target = rels.get(rel_id or "")
                if not target:
                    ctx.error("BROKEN_IMAGE_REL", f"Embedded graphic relationship {rel_id!r} is unresolved")
                    continue
                member = str((Path("word") / target).as_posix())
                try:
                    payload = archive.read(member)
                except KeyError:
                    ctx.error("MISSING_EMBEDDED_IMAGE", f"DOCX graphic member is missing: {member}")
                    continue
                records.append(
                    {
                        "payload": payload,
                        "source_name": Path(target).name,
                        "source_extension": Path(target).suffix.lower(),
                        "width_inches": image_width_inches(blip),
                        "caption_id": caption[0] if caption else "",
                        "caption": caption[1] if caption else "",
                        "paragraph_index": index,
                    }
                )
    if not records:
        ctx.error("NO_FIGURES", "No embedded Word figures were found")
        return records

    groups: dict[str, list[dict]] = {}
    unnumbered: list[str] = []
    for record in records:
        if not record["caption_id"]:
            unnumbered.append(record["source_name"])
            continue
        groups.setdefault(record["caption_id"], []).append(record)

    if unnumbered:
        ctx.note(
            "INLINE_GRAPHICS_PRESERVED",
            f"Preserved {len(unnumbered)} unnumbered inline graphic(s) in the text-only Word file; verify they are semantic table, equation, or notation content",
        )
    if not groups:
        ctx.error("NO_NUMBERED_FIGURES", "No embedded Word graphics could be associated with numbered figure captions")
        return []

    for base_id, group in groups.items():
        if len(group) > 1 and ctx.args.multipart == "combined":
            ctx.error(
                "AMBIGUOUS_COMBINATION",
                f"Figure {base_id} contains {len(group)} embedded graphics; combine them in the authoritative source rather than letting automation invent a layout",
            )
        for part_index, record in enumerate(group):
            if len(group) > 1 and ctx.args.multipart == "separate":
                record["figure_id"] = f"{base_id}{chr(ord('a') + part_index)}"
            else:
                record["figure_id"] = base_id
    return [record for record in records if record.get("figure_id")]


def image_dpi_info(image: Image.Image) -> tuple[float | None, float | None]:
    raw = image.info.get("dpi")
    if isinstance(raw, (tuple, list)) and len(raw) >= 2:
        try:
            return float(raw[0]), float(raw[1])
        except (TypeError, ValueError):
            return None, None
    return None, None


def threshold_for(figure_id: str, kind: str, args: argparse.Namespace) -> float:
    return float(args.required_dpi) if args.required_dpi else DEFAULT_THRESHOLDS[kind]


def convert_raster(
    payload: bytes,
    destination: Path,
    figure_id: str,
    width_inches: float | None,
    kind: str,
    ctx: Context,
) -> dict:
    try:
        image = Image.open(io.BytesIO(payload))
        image.load()
    except Exception as exc:
        ctx.error("UNREADABLE_RASTER", f"Figure {figure_id} could not be read: {exc}")
        return {"figure_id": figure_id, "status": "error"}

    pixels = [int(image.width), int(image.height)]
    nominal_x, nominal_y = image_dpi_info(image)
    effective_x = image.width / width_inches if width_inches and width_inches > 0 else None
    height_inches = (image.height / effective_x) if effective_x else None
    effective_y = image.height / height_inches if height_inches else effective_x
    minimum = threshold_for(figure_id, kind, ctx.args)
    if effective_x is None:
        ctx.error(
            "UNKNOWN_PRINT_WIDTH",
            f"Figure {figure_id} is raster but its printed width is unknown; supply --figure-width {figure_id}=INCHES",
        )
    elif effective_x + 0.01 < minimum:
        ctx.error(
            "LOW_EFFECTIVE_DPI",
            f"Figure {figure_id} is {effective_x:.1f} effective dpi at {width_inches:.3g} in; {kind} requires {minimum:.0f} dpi",
        )

    save_image = image
    if image.mode not in {"1", "L", "RGB", "CMYK"}:
        save_image = image.convert("RGBA" if "A" in image.getbands() else "RGB")
        if save_image.mode == "RGBA":
            background = Image.new("RGB", save_image.size, "white")
            background.paste(save_image, mask=save_image.getchannel("A"))
            save_image = background
    output_dpi = effective_x or nominal_x or minimum
    save_image.save(destination, format="TIFF", compression="tiff_lzw", dpi=(output_dpi, output_dpi))
    if destination.stat().st_size > MAX_GRAPHIC_BYTES:
        ctx.error("FIGURE_TOO_LARGE", f"{destination.name} exceeds the 15 MB ASME limit")
    if nominal_x is None:
        ctx.warn("MISSING_NOMINAL_DPI", f"Figure {figure_id} had no reliable source dpi metadata; effective dpi was evaluated from pixels and print width")

    return {
        "figure_id": figure_id,
        "file": destination.name,
        "format": "TIFF",
        "kind": kind,
        "pixels": pixels,
        "print_width_inches": round(width_inches, 5) if width_inches else None,
        "nominal_dpi": [round(nominal_x, 2), round(nominal_y, 2)] if nominal_x and nominal_y else None,
        "effective_dpi": round(effective_x, 2) if effective_x else None,
        "required_dpi": minimum,
        "bytes": destination.stat().st_size,
        "sha256": sha256(destination),
        "status": "pass" if effective_x is not None and effective_x + 0.01 >= minimum else "error",
    }


def convert_vector(source: Path, destination: Path, figure_id: str, ctx: Context) -> dict:
    if source.suffix.lower() == ".eps":
        shutil.copy2(source, destination)
    elif source.suffix.lower() == ".pdf":
        converter = shutil.which("pdftops")
        if not converter:
            ctx.error("MISSING_PDFTOPS", f"Cannot convert vector PDF for figure {figure_id} to EPS")
            return {"figure_id": figure_id, "status": "error"}
        result = subprocess.run(
            [converter, "-eps", str(source), str(destination)],
            capture_output=True,
            text=True,
            check=False,
        )
        if result.returncode != 0 or not destination.exists():
            ctx.error("VECTOR_CONVERSION_FAILED", f"Figure {figure_id} PDF-to-EPS conversion failed: {result.stderr.strip()}")
            return {"figure_id": figure_id, "status": "error"}
    if destination.stat().st_size > MAX_GRAPHIC_BYTES:
        ctx.error("FIGURE_TOO_LARGE", f"{destination.name} exceeds the 15 MB ASME limit")
    return {
        "figure_id": figure_id,
        "file": destination.name,
        "format": "EPS",
        "kind": "vector",
        "pixels": None,
        "print_width_inches": None,
        "nominal_dpi": None,
        "effective_dpi": None,
        "required_dpi": None,
        "bytes": destination.stat().st_size,
        "sha256": sha256(destination),
        "status": "pass",
    }


def write_payload_file(payload: bytes, suffix: str) -> Path:
    handle = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
    try:
        handle.write(payload)
        return Path(handle.name)
    finally:
        handle.close()


def process_figure_records(records: list[dict], ctx: Context, kinds: dict[str, str], widths: dict[str, float]) -> None:
    figure_dir = ctx.upload / "figures"
    figure_dir.mkdir(parents=True, exist_ok=True)
    seen: set[str] = set()
    for record in sorted(records, key=lambda item: figure_sort_key(item["figure_id"])):
        figure_id = record["figure_id"]
        if figure_id in seen:
            ctx.error("DUPLICATE_FIGURE_ID", f"More than one source graphic maps to Figure {figure_id}")
            continue
        seen.add(figure_id)
        suffix = record["source_extension"]
        width = widths.get(figure_id, record.get("width_inches"))
        if suffix in RASTER_EXTENSIONS:
            destination = figure_dir / f"Fig_{figure_id}.tiff"
            kind = kinds.get(figure_id, kinds.get(re.match(r"\d+", figure_id).group(), "linework"))
            result = convert_raster(record["payload"], destination, figure_id, width, kind, ctx)
        elif suffix in VECTOR_EXTENSIONS:
            destination = figure_dir / f"Fig_{figure_id}.eps"
            temp = write_payload_file(record["payload"], suffix)
            try:
                result = convert_vector(temp, destination, figure_id, ctx)
            finally:
                temp.unlink(missing_ok=True)
        else:
            ctx.error(
                "UNSUPPORTED_GRAPHIC",
                f"Figure {figure_id} uses {suffix or 'an unknown format'}; provide TIFF, EPS, or vector PDF source",
            )
            result = {"figure_id": figure_id, "status": "error"}
        result["source_name"] = record.get("source_name")
        result["caption"] = record.get("caption", "")
        ctx.figures.append(result)


def find_soffice() -> str | None:
    return shutil.which("soffice") or shutil.which("libreoffice")


def convert_docx_to_pdf(source: Path, destination: Path, ctx: Context, code: str) -> bool:
    soffice = find_soffice()
    if not soffice:
        ctx.error(code, "LibreOffice/soffice is required to render Word source")
        return False
    destination.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="jmd-soffice-") as tmp:
        result = subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp, str(source)],
            capture_output=True,
            text=True,
            check=False,
        )
        generated = Path(tmp) / f"{source.stem}.pdf"
        if result.returncode != 0 or not generated.exists():
            ctx.error(code, f"Word-to-PDF rendering failed: {(result.stderr or result.stdout).strip()}")
            return False
        shutil.copy2(generated, destination)
    return validate_pdf(destination, ctx, code)


def validate_pdf(path: Path, ctx: Context, code: str = "INVALID_PDF") -> bool:
    try:
        reader = PdfReader(str(path))
        if not reader.pages:
            raise ValueError("PDF has no pages")
    except Exception as exc:
        ctx.error(code, f"{path.name} is not a readable nonempty PDF: {exc}")
        return False
    return True


def check_pdf_blank_pages(path: Path, ctx: Context, label: str) -> None:
    """Block pages that contain neither meaningful text nor placed artwork."""
    try:
        reader = PdfReader(str(path))
    except Exception:
        return
    blank_pages: list[int] = []
    graphics_only_pages: list[int] = []
    for page_number, page in enumerate(reader.pages, 1):
        try:
            text = page.extract_text() or ""
        except Exception:
            text = ""
        words = re.findall(r"[^\W\d_]{2,}", text, flags=re.UNICODE)
        if len(words) >= 2:
            continue
        contents = page.get_contents()
        content_bytes = contents.get_data() if contents is not None else b""
        placed_artwork = bool(re.search(rb"/[^\s]+\s+Do\b", content_bytes))
        if placed_artwork:
            graphics_only_pages.append(page_number)
        else:
            blank_pages.append(page_number)
    if blank_pages:
        pages = ", ".join(str(value) for value in blank_pages)
        ctx.error(
            "BLANK_OR_NEAR_BLANK_PAGE",
            f"{label} contains page(s) with no meaningful text or placed artwork: {pages}",
        )
    if graphics_only_pages:
        pages = ", ".join(str(value) for value in graphics_only_pages)
        ctx.warn(
            "GRAPHICS_ONLY_PAGE",
            f"Inspect graphics-only page(s) in {label} manually: {pages}",
        )


def render_pdf(path: Path, target: Path, ctx: Context, label: str) -> None:
    renderer = shutil.which("pdftoppm")
    if not renderer:
        ctx.warn("MISSING_PDF_RENDERER", f"pdftoppm is unavailable; {label} pages were not rendered for QA")
        return
    target.mkdir(parents=True, exist_ok=True)
    prefix = target / "page"
    result = subprocess.run(
        [renderer, "-png", "-r", "150", str(path), str(prefix)],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        ctx.error("PDF_RENDER_FAILED", f"Could not render {label} for QA: {result.stderr.strip()}")


def check_pdf_fonts(path: Path, ctx: Context) -> None:
    tool = shutil.which("pdffonts")
    if not tool:
        ctx.warn("MISSING_FONT_TOOL", "pdffonts is unavailable; embedded-font status was not checked")
        return
    result = subprocess.run([tool, str(path)], capture_output=True, text=True, check=False)
    if result.returncode != 0:
        ctx.warn("FONT_CHECK_FAILED", f"Could not inspect fonts in {path.name}")
        return
    rows = [line.split() for line in result.stdout.splitlines()[2:] if line.strip()]
    if rows and any(len(row) >= 5 and row[4].lower() == "no" for row in rows):
        ctx.warn("UNEMBEDDED_FONT", f"{path.name} appears to contain at least one unembedded font")


def check_required_sections(text: str, ctx: Context) -> None:
    lowered = re.sub(r"\s+", " ", text).lower()
    for required in ctx.args.require_section:
        if re.sub(r"\s+", " ", required).lower() not in lowered:
            ctx.error("MISSING_REQUIRED_SECTION", f"Required text/heading not found: {required}")
    for heading in ("data availability", "conflict of interest"):
        if heading not in lowered:
            ctx.warn("REVIEW_REQUIRED_STATEMENT", f"Verify whether a {heading.title()} statement is required")


def caption_and_callout_sets(text: str) -> tuple[set[str], set[str]]:
    captions: set[str] = set()
    callouts: set[str] = set()
    for line in re.split(r"[\r\n]+", text):
        caption = FIGURE_CAPTION_RE.match(line)
        if caption:
            captions.add(normalize_figure_id("".join(caption.groups(default=""))))
        else:
            for match in FIGURE_CALLOUT_RE.finditer(line):
                callouts.add(normalize_figure_id("".join(match.groups(default=""))))
            for group in FIGURE_GROUP_CALLOUT_RE.findall(line):
                callouts.update(
                    normalize_figure_id(value)
                    for value in re.findall(r"\d+\s*[a-z]?", group, flags=re.IGNORECASE)
                )
    return captions, callouts


def process_docx(source: Path, ctx: Context, kinds: dict[str, str], widths: dict[str, float]) -> None:
    paper = safe_name(ctx.args.paper_id)
    clean_docx = ctx.root / "work" / f"{paper}_clean.docx"
    text_only = ctx.upload / f"{paper}_text_only.docx"
    clean_docx.parent.mkdir(parents=True, exist_ok=True)
    ctx.upload.mkdir(parents=True, exist_ok=True)
    clean_counts = sanitize_docx(source, clean_docx, remove_graphics=False)
    text_counts = remove_numbered_docx_figures(clean_docx, text_only)
    changed = sum(clean_counts.values())
    if changed:
        ctx.note(
            "WORD_MARKUP_RESOLVED",
            f"Accepted {clean_counts['insertions']} insertion group(s), removed "
            f"{clean_counts['deletions']} deletion group(s) and "
            f"{clean_counts['deleted_paragraphs']} emptied deleted paragraph(s), removed "
            f"{clean_counts['format_changes']} formatting-revision record(s), and removed "
            f"{clean_counts['comments']} comment marker(s)",
        )

    clean_text = docx_visible_text(clean_docx)
    text_only_text = docx_visible_text(text_only)
    ctx.source_text = clean_text
    if clean_text != text_only_text:
        ctx.error("TEXT_ONLY_MISMATCH", "Removing graphics changed visible Word text; figure callouts or captions may have been lost")
    if docx_has_numbered_figures(text_only):
        ctx.error("NUMBERED_FIGURES_REMAIN", "The generated Word text-only file still contains numbered production figures")
    if text_counts["graphics"] == 0:
        ctx.warn("NO_GRAPHICS_REMOVED", "No numbered Word drawing/object nodes were removed from the text-only file")

    records = extract_docx_graphics(clean_docx, ctx)
    process_figure_records(records, ctx, kinds, widths)
    captions, callouts = caption_and_callout_sets(clean_text)
    uploaded = {item["figure_id"] for item in ctx.figures if item.get("file")}
    uploaded_bases = {re.match(r"\d+", value).group() for value in uploaded}
    caption_bases = {re.match(r"\d+", value).group() for value in captions}
    if caption_bases and uploaded_bases != caption_bases:
        ctx.error(
            "FIGURE_SET_MISMATCH",
            f"Embedded/uploaded figure numbers {sorted(uploaded_bases, key=int)} do not match caption numbers {sorted(caption_bases, key=int)}",
        )
    if caption_bases and not caption_bases.issubset({re.match(r"\d+", value).group() for value in callouts}):
        missing = caption_bases - {re.match(r"\d+", value).group() for value in callouts}
        ctx.warn("UNCALLED_FIGURE", f"Verify in-text callouts for figure(s): {', '.join(sorted(missing, key=int))}")

    complete_pdf = ctx.upload / f"{paper}_complete.pdf"
    if convert_docx_to_pdf(clean_docx, complete_pdf, ctx, "COMPLETE_PDF_FAILED"):
        ctx.complete_pdf = complete_pdf
        ctx.warn(
            "WORD_RENDERER_VARIANCE",
            "The complete PDF was rendered with LibreOffice; compare pagination, fonts, equations, and object placement with Microsoft Word before upload",
        )
        check_pdf_fonts(complete_pdf, ctx)
        check_pdf_blank_pages(complete_pdf, ctx, "complete PDF")
        render_pdf(complete_pdf, ctx.qa / "complete", ctx, "complete PDF")
    qa_text_pdf = ctx.root / "work" / f"{paper}_text_only.pdf"
    if convert_docx_to_pdf(text_only, qa_text_pdf, ctx, "TEXT_ONLY_RENDER_FAILED"):
        ctx.text_only_render = qa_text_pdf
        check_pdf_blank_pages(qa_text_pdf, ctx, "rendered text-only Word file")
        render_pdf(qa_text_pdf, ctx.qa / "text-only", ctx, "text-only Word file")


def strip_tex_comments(text: str) -> str:
    return re.sub(r"(?<!\\)%.*", "", text)


def find_tex_main(project: Path, requested: str | None, ctx: Context) -> Path | None:
    if requested:
        candidate = (project / requested).resolve()
        if project.resolve() in candidate.parents and candidate.is_file():
            return candidate
        requested_parts = Path(requested).parts
        matches = [
            path
            for path in project.rglob(Path(requested).name)
            if tuple(path.relative_to(project).parts[-len(requested_parts) :]) == requested_parts
        ]
        if len(matches) == 1:
            return matches[0]
        ctx.error("MISSING_MAIN_TEX", f"Requested root TeX file does not exist uniquely: {requested}")
        return None
    candidates = []
    for path in project.rglob("*.tex"):
        try:
            text = path.read_text(encoding="utf-8", errors="replace")
        except OSError:
            continue
        if "\\documentclass" in text and "\\begin{document}" in text:
            candidates.append(path)
    if len(candidates) == 1:
        return candidates[0]
    if not candidates:
        ctx.error("MISSING_MAIN_TEX", "No root TeX file with documentclass and begin{document} was found")
    else:
        ctx.error("AMBIGUOUS_MAIN_TEX", "Multiple root TeX candidates found; supply --main")
    return None


def resolve_tex_include(current: Path, raw: str) -> Path | None:
    candidate = (current.parent / raw).resolve()
    if candidate.suffix == "":
        candidate = candidate.with_suffix(".tex")
    return candidate if candidate.is_file() else None


def expand_tex(path: Path, seen: set[Path] | None = None) -> str:
    seen = seen or set()
    resolved = path.resolve()
    if resolved in seen:
        return ""
    seen.add(resolved)
    text = path.read_text(encoding="utf-8", errors="replace")

    def replace(match: re.Match[str]) -> str:
        included = resolve_tex_include(path, match.group(1).strip())
        return expand_tex(included, seen) if included else match.group(0)

    return re.sub(r"\\(?:input|include)\s*\{([^}]+)\}", replace, text)


def bib_keys(path: Path) -> set[str]:
    text = path.read_text(encoding="utf-8", errors="replace")
    return set(re.findall(r"@[A-Za-z]+\s*\{\s*([^,\s]+)\s*,", strip_tex_comments(text)))


def citation_audit(tex_text: str, bibs: list[Path], ctx: Context) -> None:
    clean = strip_tex_comments(tex_text)
    if re.search(r"\\nocite\s*\{\s*\*\s*\}", clean):
        ctx.error("NOCITE_STAR", r"\nocite{*} is present and can push uncited bibliography entries into production")
    cited: set[str] = set()
    for match in re.finditer(r"\\(?:cite\w*|nocite)\s*(?:\[[^\]]*\]\s*)*\{([^}]+)\}", clean):
        cited.update(key.strip() for key in match.group(1).split(",") if key.strip() != "*")
    available: set[str] = set()
    for bib in bibs:
        available.update(bib_keys(bib))
    missing = cited - available
    unused = available - cited
    if missing:
        ctx.error("MISSING_BIB_KEY", f"Citation key(s) absent from uploaded .bib files: {', '.join(sorted(missing))}")
    if unused:
        ctx.note("UNUSED_BIB_KEYS", f"{len(unused)} .bib record(s) are not cited; this is safe unless the source forces them into the reference list")
    if not cited:
        ctx.warn("NO_CITATIONS_DETECTED", "No LaTeX citation commands were detected; audit references manually")


def dimension_inches(value: str) -> float | None:
    match = re.fullmatch(r"\s*([0-9]*\.?[0-9]+)\s*(in|cm|mm|pt)\s*", value)
    if not match:
        return None
    amount = float(match.group(1))
    return amount * {"in": 1.0, "cm": 1 / 2.54, "mm": 1 / 25.4, "pt": 1 / 72.27}[match.group(2)]


def relative_width_inches(value: str, scope: dict[str, float | None]) -> float | None:
    absolute = dimension_inches(value)
    if absolute is not None:
        return absolute
    match = re.fullmatch(
        r"\s*([0-9]*\.?[0-9]+)?\s*\\(linewidth|textwidth|columnwidth)\s*",
        value,
    )
    if not match:
        return None
    base = scope.get(match.group(2))
    if base is None:
        return None
    return float(match.group(1) or 1.0) * base


def parse_width_option(options: str, scope: dict[str, float | None] | None = None) -> float | None:
    match = re.search(
        r"\bwidth\s*=\s*([^,\]]+)",
        options,
    )
    if not match:
        return None
    return relative_width_inches(match.group(1), scope or {})


def infer_latex_layout_widths(
    project: Path, tex_text: str, ctx: Context
) -> dict[str, float | None]:
    text_width = ctx.args.text_width
    column_width = ctx.args.column_width
    inferred: list[str] = []

    class_match = re.search(
        r"\\documentclass\s*(?:\[[^\]]*\])?\s*\{([^}]+)\}",
        strip_tex_comments(tex_text),
    )
    class_text = ""
    if class_match:
        class_name = class_match.group(1).strip()
        candidates = [project / f"{class_name}.cls", *project.rglob(f"{Path(class_name).name}.cls")]
        class_file = next((path for path in candidates if path.is_file()), None)
        if class_file:
            class_text = strip_tex_comments(
                class_file.read_text(encoding="utf-8", errors="replace")
            )

    geometry: dict[str, str] = {}
    if class_text:
        geometry_match = re.search(
            r"\\(?:RequirePackage|usepackage)\s*\[([^\]]+)\]\s*\{geometry\}",
            class_text,
            flags=re.DOTALL,
        )
        if geometry_match:
            for item in geometry_match.group(1).split(","):
                if "=" in item:
                    key, value = item.split("=", 1)
                    geometry[key.strip().lower()] = value.strip()

    if text_width is None:
        text_width = dimension_inches(geometry.get("textwidth", ""))
        if text_width is None:
            paper_width = dimension_inches(geometry.get("paperwidth", ""))
            left = dimension_inches(geometry.get("left", ""))
            right = dimension_inches(geometry.get("right", ""))
            if paper_width is not None and left is not None and right is not None:
                text_width = paper_width - left - right
        if text_width is not None:
            inferred.append(f"text width {text_width:.3f} in")

    if column_width is None and text_width is not None:
        column_sep = None
        if class_text:
            separator_match = re.search(
                r"\\setlength\s*\{?\\columnsep\}?\s*\{([^}]+)\}",
                class_text,
            )
            if separator_match:
                column_sep = dimension_inches(separator_match.group(1))
        if "twocolumn" in class_text and column_sep is not None:
            column_width = (text_width - column_sep) / 2
        elif "twocolumn" not in class_text:
            column_width = text_width
        if column_width is not None:
            inferred.append(f"column width {column_width:.3f} in")

    if inferred:
        ctx.note("LATEX_LAYOUT_INFERRED", "Inferred " + " and ".join(inferred) + " from the local document class")
    return {"textwidth": text_width, "columnwidth": column_width}


def scoped_latex_graphic_width(
    environment: str,
    graphic_offset: int,
    options: str,
    figure_starred: bool,
    layout: dict[str, float | None],
) -> float | None:
    base = {
        "textwidth": layout.get("textwidth"),
        "columnwidth": layout.get("columnwidth"),
        "linewidth": layout.get("textwidth") if figure_starred else layout.get("columnwidth"),
    }
    stack: list[tuple[str, dict[str, float | None]]] = [("figure", base)]
    token_re = re.compile(
        r"\\begin\{(minipage|subfigure)\}(?:\[[^\]]*\])?\s*\{([^}]+)\}"
        r"|\\end\{(minipage|subfigure)\}"
    )
    for token in token_re.finditer(environment[:graphic_offset]):
        if token.group(1):
            width = relative_width_inches(token.group(2), stack[-1][1])
            nested = (
                {"textwidth": width, "columnwidth": width, "linewidth": width}
                if width is not None
                else stack[-1][1].copy()
            )
            stack.append((token.group(1), nested))
        else:
            closing = token.group(3)
            while len(stack) > 1:
                name, _ = stack.pop()
                if name == closing:
                    break
    return parse_width_option(options, stack[-1][1])


def resolve_graphic(project: Path, raw: str, ctx: Context) -> Path | None:
    cleaned = raw.strip().replace("\\detokenize{", "").rstrip("}")
    direct = project / cleaned
    candidates: list[Path] = []
    if direct.suffix:
        if direct.is_file():
            candidates.append(direct)
    else:
        for extension in GRAPHIC_EXTENSIONS:
            candidate = direct.with_suffix(extension)
            if candidate.is_file():
                candidates.append(candidate)
    if not candidates:
        name = Path(cleaned).name
        for path in project.rglob("*"):
            if not path.is_file():
                continue
            if Path(cleaned).suffix:
                if path.name == name:
                    candidates.append(path)
            elif path.stem == name and path.suffix.lower() in GRAPHIC_EXTENSIONS:
                candidates.append(path)
    unique = sorted({path.resolve() for path in candidates})
    if len(unique) == 1:
        return unique[0]
    if not unique:
        ctx.error("MISSING_GRAPHIC", f"LaTeX graphic could not be resolved: {raw}")
    else:
        ctx.error("AMBIGUOUS_GRAPHIC", f"LaTeX graphic resolves to multiple files: {raw}")
    return None


def latex_figure_records(
    project: Path,
    tex_text: str,
    ctx: Context,
    layout: dict[str, float | None],
) -> list[dict]:
    clean = strip_tex_comments(tex_text)
    environments = list(
        re.finditer(
            r"\\begin\{(figure\*?)\}(.*?)\\end\{\1\}", clean, flags=re.DOTALL
        )
    )
    records: list[dict] = []
    for number, environment_match in enumerate(environments, start=1):
        environment_name = environment_match.group(1)
        environment = environment_match.group(2)
        graphics = list(
            re.finditer(
                r"\\includegraphics\s*(?:\[([^\]]*)\])?\s*\{([^}]+)\}",
                environment,
            )
        )
        caption_match = re.search(r"\\caption(?:\[[^\]]*\])?\s*\{(.*?)\}", environment, flags=re.DOTALL)
        caption = re.sub(r"\s+", " ", caption_match.group(1)).strip() if caption_match else ""
        labels = re.findall(r"\\label\s*\{([^}]+)\}", environment)
        if not graphics:
            ctx.warn("FIGURE_WITHOUT_GRAPHIC", f"LaTeX figure environment {number} has no includegraphics command")
            continue
        if not caption:
            ctx.error("MISSING_FIGURE_CAPTION", f"LaTeX figure environment {number} has no caption")
        if len(graphics) > 1 and ctx.args.multipart == "combined":
            ctx.error(
                "AMBIGUOUS_COMBINATION",
                f"Figure {number} contains {len(graphics)} graphics; combine them in source rather than inventing an automated layout",
            )
        for index, graphic in enumerate(graphics):
            options, raw_path = graphic.group(1), graphic.group(2)
            source = resolve_graphic(project, raw_path, ctx)
            if source is None:
                continue
            figure_id = f"{number}{chr(ord('a') + index)}" if len(graphics) > 1 and ctx.args.multipart == "separate" else str(number)
            records.append(
                {
                    "payload": source.read_bytes(),
                    "source_name": source.name,
                    "source_extension": source.suffix.lower(),
                    "width_inches": scoped_latex_graphic_width(
                        environment,
                        graphic.start(),
                        options or "",
                        environment_name.endswith("*"),
                        layout,
                    ),
                    "figure_id": figure_id,
                    "caption": caption,
                    "labels": labels,
                }
            )
    if not environments:
        ctx.error("NO_FIGURES", "No LaTeX figure environments were found")
    defined_labels = {label for record in records for label in record.get("labels", [])}
    referenced_labels = set(re.findall(r"\\(?:ref|autoref|cref|Cref)\s*\{([^}]+)\}", clean))
    uncalled = defined_labels - referenced_labels
    if uncalled:
        ctx.warn("UNCALLED_FIGURE_LABEL", f"Figure label(s) not referenced in source: {', '.join(sorted(uncalled))}")
    return records


def resolve_local_latex_dependency(
    project: Path, current: Path, raw: str, suffix: str
) -> Path | None:
    cleaned = raw.strip().strip('"\'')
    candidate = current.parent / cleaned
    if not candidate.suffix:
        candidate = candidate.with_suffix(suffix)
    if candidate.is_file():
        return candidate.resolve()
    candidate = project / cleaned
    if not candidate.suffix:
        candidate = candidate.with_suffix(suffix)
    if candidate.is_file():
        return candidate.resolve()
    name = Path(cleaned).name
    if not Path(name).suffix:
        name += suffix
    matches = [path.resolve() for path in project.rglob(name) if path.is_file()]
    return matches[0] if len(matches) == 1 else None


def latex_dependency_files(project: Path, main: Path, ctx: Context) -> set[Path]:
    required: set[Path] = {main.resolve()}
    pending = [main.resolve()]
    patterns = [
        (r"\\(?:input|include)\s*\{([^}]+)\}", ".tex", True),
        (r"\\bibliography\s*\{([^}]+)\}", ".bib", True),
        (r"\\addbibresource(?:\[[^\]]*\])?\s*\{([^}]+)\}", ".bib", True),
        (r"\\documentclass\s*(?:\[[^\]]*\])?\s*\{([^}]+)\}", ".cls", False),
        (r"\\usepackage\s*(?:\[[^\]]*\])?\s*\{([^}]+)\}", ".sty", False),
        (r"\\bibliographystyle\s*\{([^}]+)\}", ".bst", False),
    ]
    while pending:
        current = pending.pop()
        if current.suffix.lower() not in {".tex", ".cls", ".sty"}:
            continue
        text = strip_tex_comments(current.read_text(encoding="utf-8", errors="replace"))
        for pattern, suffix, must_exist in patterns:
            for match in re.finditer(pattern, text):
                for raw in match.group(1).split(","):
                    dependency = resolve_local_latex_dependency(project, current, raw, suffix)
                    if dependency is None:
                        if must_exist:
                            ctx.error(
                                "MISSING_LATEX_DEPENDENCY",
                                f"Could not resolve {raw.strip()}{suffix if not Path(raw.strip()).suffix else ''} referenced by {current.relative_to(project)}",
                            )
                        continue
                    if dependency not in required:
                        required.add(dependency)
                        pending.append(dependency)
    return required


def copy_latex_text_only(
    project: Path,
    main: Path,
    destination: Path,
    dependencies: set[Path],
    ctx: Context,
) -> None:
    destination.mkdir(parents=True, exist_ok=True)
    for source in sorted(dependencies):
        relative = source.relative_to(project)
        target = destination / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        if source.resolve() == main.resolve():
            text = source.read_text(encoding="utf-8", errors="replace")
            marker = "\\begin{document}"
            if marker not in text:
                ctx.error("MISSING_BEGIN_DOCUMENT", f"Root TeX file {relative} has no begin{{document}}")
                continue
            replacement = marker + "\n% Production text-only override inserted by prepare-asme-submission-package\n\\renewcommand{\\includegraphics}[2][]{\\relax}\n"
            target.write_text(text.replace(marker, replacement, 1), encoding="utf-8")
        else:
            shutil.copy2(source, target)
    eligible = {
        source.resolve()
        for source in project.rglob("*")
        if source.is_file()
        and source.suffix.lower() not in GRAPHIC_EXTENSIONS
        and source.suffix.lower() not in {".aux", ".log", ".out", ".blg", ".bbl", ".synctex.gz"}
    }
    pruned = eligible - dependencies
    if pruned:
        ctx.note(
            "LATEX_UNUSED_FILES_PRUNED",
            f"Excluded {len(pruned)} unreferenced non-graphic project file(s), including archived manuscript versions, from the upload source set",
        )


def compile_latex(project: Path, main: Path, destination: Path, ctx: Context) -> bool:
    latexmk = shutil.which("latexmk")
    if not latexmk:
        ctx.error("MISSING_FULL_PDF", "Supply --full-pdf because latexmk is not installed")
        return False
    result = subprocess.run(
        [latexmk, "-pdf", "-interaction=nonstopmode", "-halt-on-error", main.name],
        cwd=main.parent,
        capture_output=True,
        text=True,
        check=False,
    )
    generated = main.with_suffix(".pdf")
    if result.returncode != 0 or not generated.exists():
        ctx.error("LATEX_COMPILE_FAILED", f"LaTeX compilation failed: {(result.stdout + result.stderr)[-3000:]}")
        return False
    shutil.copy2(generated, destination)
    return validate_pdf(destination, ctx)


def rendered_reference_audit(path: Path, ctx: Context) -> None:
    try:
        text = "\n".join(page.extract_text() or "" for page in PdfReader(str(path)).pages)
    except Exception:
        return
    headings = list(re.finditer(r"(?im)^\s*references\s*\d*\s*$", text))
    if not headings:
        ctx.warn("NO_RENDERED_REFERENCES_HEADING", "Could not locate a References heading in the complete PDF")
        return
    tail = text[headings[-1].end() :]
    numbers = [int(value) for value in re.findall(r"(?m)^\s*\[(\d+)\]\s+", tail)]
    if len(numbers) >= 3:
        unique = list(dict.fromkeys(numbers))
        expected = list(range(unique[0], unique[-1] + 1))
        if unique != expected or unique[0] != 1:
            ctx.warn("REFERENCE_NUMBER_GAPS", f"Rendered reference numbering may contain gaps: {unique}")
    else:
        ctx.warn("REFERENCE_AUDIT_MANUAL", "Rendered reference numbering could not be parsed reliably; compare citations and references manually")


def process_latex_zip(source: Path, ctx: Context, kinds: dict[str, str], widths: dict[str, float]) -> None:
    project = ctx.root / "work" / "overleaf"
    project.mkdir(parents=True, exist_ok=True)
    if not safe_extract_zip(source, project, ctx):
        return
    main = find_tex_main(project, ctx.args.main, ctx)
    if main is None:
        return
    ctx.latex_main_relative = main.relative_to(project)
    tex_text = expand_tex(main)
    ctx.source_text = tex_text
    dependencies = latex_dependency_files(project, main, ctx)
    bibs = sorted(path for path in dependencies if path.suffix.lower() == ".bib")
    if not bibs:
        ctx.error("MISSING_BIB", "No separate .bib file is referenced by the root LaTeX source")
    else:
        citation_audit(tex_text, bibs, ctx)
    layout = infer_latex_layout_widths(project, tex_text, ctx)
    records = latex_figure_records(project, tex_text, ctx, layout)
    process_figure_records(records, ctx, kinds, widths)
    native = ctx.upload / "native"
    copy_latex_text_only(project, main, native, dependencies, ctx)

    latexmk = shutil.which("latexmk")
    if latexmk:
        compile_tree = ctx.root / "work" / "text-only-compile"
        shutil.copytree(native, compile_tree)
        compile_main = compile_tree / main.relative_to(project)
        text_only_pdf = ctx.root / "work" / f"{safe_name(ctx.args.paper_id)}_text_only.pdf"
        if compile_latex(compile_tree, compile_main, text_only_pdf, ctx):
            ctx.text_only_render = text_only_pdf
            check_pdf_blank_pages(text_only_pdf, ctx, "rendered text-only LaTeX source")
            render_pdf(text_only_pdf, ctx.qa / "text-only", ctx, "text-only LaTeX source")
    else:
        ctx.warn(
            "TEXT_ONLY_NOT_RENDERED",
            "latexmk is unavailable, so the generated text-only LaTeX source must be compiled and inspected before upload",
        )

    paper = safe_name(ctx.args.paper_id)
    complete_pdf = ctx.upload / f"{paper}_complete.pdf"
    if ctx.args.full_pdf:
        if not ctx.args.full_pdf.is_file():
            ctx.error("MISSING_FULL_PDF", f"Complete PDF does not exist: {ctx.args.full_pdf}")
        else:
            shutil.copy2(ctx.args.full_pdf, complete_pdf)
            if validate_pdf(complete_pdf, ctx):
                ctx.complete_pdf = complete_pdf
                ctx.warn(
                    "SUPPLIED_FULL_PDF",
                    "The complete PDF was supplied separately; verify it was compiled from the exact source revision in this package",
                )
    elif compile_latex(project, main, complete_pdf, ctx):
        ctx.complete_pdf = complete_pdf
    if ctx.complete_pdf:
        check_pdf_fonts(ctx.complete_pdf, ctx)
        check_pdf_blank_pages(ctx.complete_pdf, ctx, "complete PDF")
        rendered_reference_audit(ctx.complete_pdf, ctx)
        render_pdf(ctx.complete_pdf, ctx.qa / "complete", ctx, "complete PDF")


def check_text_only_native(ctx: Context) -> None:
    if ctx.args.source.suffix.lower() == ".docx":
        return
    native = ctx.upload / "native"
    if not list(native.rglob("*.bib")):
        ctx.error("BIB_NOT_UPLOADED", "The generated native source set does not contain a separate .bib file")
    for tex in native.rglob("*.tex"):
        text = tex.read_text(encoding="utf-8", errors="replace")
        relative = tex.relative_to(native)
        if relative == ctx.latex_main_relative or "\\documentclass" in text:
            if "Production text-only override" not in text:
                ctx.error("LATEX_GRAPHICS_NOT_DISABLED", f"Graphics were not disabled in root TeX file {relative}")


def copy_supplemental_files(ctx: Context) -> None:
    if not ctx.args.supplemental_file:
        return
    destination = ctx.upload / "supplemental"
    destination.mkdir(parents=True, exist_ok=True)
    seen: set[str] = set()
    for source in ctx.args.supplemental_file:
        if not source.is_file():
            ctx.error("MISSING_SUPPLEMENTAL_FILE", f"Supplemental file does not exist: {source}")
            continue
        if source.name in seen or (destination / source.name).exists():
            ctx.error("DUPLICATE_SUPPLEMENTAL_NAME", f"Supplemental filename is not unique: {source.name}")
            continue
        seen.add(source.name)
        shutil.copy2(source, destination / source.name)


def package_files(upload: Path, destination: Path) -> None:
    with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(upload.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(upload))


def output_inventory(root: Path, excluded: set[Path] | None = None) -> list[dict]:
    excluded = {path.resolve() for path in (excluded or set())}
    items = []
    for path in sorted(root.rglob("*")):
        if path.is_file() and path.resolve() not in excluded:
            items.append(
                {
                    "path": str(path.relative_to(root)),
                    "bytes": path.stat().st_size,
                    "sha256": sha256(path),
                }
            )
    return items


def write_report(ctx: Context) -> None:
    has_warnings = any(finding.level == "WARNING" for finding in ctx.findings)
    status = "BLOCKED" if ctx.blocked else ("PASS — MANUAL WARNINGS REMAIN" if has_warnings else "PASS")
    raster_rule = (
        f"paper-specific {ctx.args.required_dpi:g} dpi minimum"
        if ctx.args.required_dpi
        else "linework at 900 dpi unless explicitly classified"
    )
    lines = [
        f"# QA report: {ctx.args.paper_id}",
        "",
        f"**Status:** {status}",
        "",
        f"- Journal: {ctx.args.journal.upper()}",
        f"- Source: `{ctx.args.source.name}`",
        f"- Due date: {ctx.args.due_date or 'not recorded'}",
        f"- Multipart rule: `{ctx.args.multipart}`",
        f"- Raster rule: {raster_rule}",
        "",
    ]
    for level in ("ERROR", "WARNING", "NOTE"):
        findings = [finding for finding in ctx.findings if finding.level == level]
        lines.extend([f"## {level.title()}s", ""])
        if findings:
            lines.extend(f"- `{finding.code}` — {finding.message}" for finding in findings)
        else:
            lines.append("- None.")
        lines.append("")
    lines.extend(
        [
            "## Figures",
            "",
            "| Figure | File | Kind | Effective dpi | Required dpi | Status |",
            "|---|---|---:|---:|---:|---|",
        ]
    )
    for figure in sorted(ctx.figures, key=lambda item: figure_sort_key(item.get("figure_id", ""))):
        lines.append(
            "| {figure_id} | {file} | {kind} | {effective} | {required} | {status} |".format(
                figure_id=figure.get("figure_id", "?"),
                file=figure.get("file", "—"),
                kind=figure.get("kind", "—"),
                effective=figure.get("effective_dpi", "vector/unknown"),
                required=figure.get("required_dpi", "—"),
                status=figure.get("status", "error"),
            )
        )
    lines.extend(
        [
            "",
            "## Manual closeout",
            "",
            "- Inspect every page under `qa/complete/` and `qa/text-only/` when present.",
            "- Inspect every TIFF/EPS at intended publication size.",
            "- For unworkable figures, search only authorized source locations or ask the user once for permission; otherwise request native/high-resolution replacements with exact specifications.",
            "- Confirm caption lists, callouts, tables, equations, links, statements, and page footers.",
            "- Confirm the acceptance letter's exact multipart, resolution, and portal naming instructions.",
            "- Rebuild the full package after any change.",
            "",
        ]
    )
    (ctx.root / "QA_REPORT.md").write_text("\n".join(lines), encoding="utf-8")


def write_manifest(ctx: Context, zip_path: Path | None) -> None:
    manifest_path = ctx.root / "manifest.json"
    inventory = output_inventory(ctx.root, {manifest_path})
    manifest = {
        "schema_version": 1,
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "paper_id": ctx.args.paper_id,
        "journal": ctx.args.journal,
        "source": {
            "name": ctx.args.source.name,
            "sha256": sha256(ctx.args.source),
        },
        "controlling_requirements": {
            "multipart": ctx.args.multipart,
            "required_dpi_override": ctx.args.required_dpi,
            "required_sections": ctx.args.require_section,
            "due_date": ctx.args.due_date,
            "editorial_instructions": ctx.args.instruction,
            "portal_rules": ctx.args.portal_rule,
            "supplemental_files": [path.name for path in ctx.args.supplemental_file],
            "default_raster_kind": "linework",
            "thresholds": DEFAULT_THRESHOLDS,
        },
        "status": "blocked" if ctx.blocked else "pass",
        "manual_review_required": any(finding.level == "WARNING" for finding in ctx.findings),
        "findings": [finding.__dict__ for finding in ctx.findings],
        "figures": ctx.figures,
        "upload_zip": zip_path.name if zip_path else None,
        "outputs": inventory,
    }
    manifest_path.write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")


def finalize(ctx: Context) -> int:
    check_required_sections(ctx.source_text, ctx)
    check_text_only_native(ctx)
    copy_supplemental_files(ctx)
    zip_path: Path | None = None
    if not ctx.blocked:
        zip_path = ctx.root / f"{safe_name(ctx.args.paper_id)}_submission_package.zip"
        package_files(ctx.upload, zip_path)
    write_report(ctx)
    write_manifest(ctx, zip_path)
    return 2 if ctx.blocked else 0


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    args.source = args.source.expanduser().resolve()
    args.output = args.output.expanduser().resolve()
    if args.full_pdf:
        args.full_pdf = args.full_pdf.expanduser().resolve()
    args.supplemental_file = [path.expanduser().resolve() for path in args.supplemental_file]
    if not args.source.is_file():
        print(f"Source does not exist: {args.source}", file=sys.stderr)
        return 2
    if args.output.exists():
        print(f"Output already exists; choose a new path or remove it intentionally: {args.output}", file=sys.stderr)
        return 2
    if args.required_dpi is not None and args.required_dpi <= 0:
        print("--required-dpi must be positive", file=sys.stderr)
        return 2
    if args.text_width is not None and args.text_width <= 0:
        print("--text-width must be positive", file=sys.stderr)
        return 2
    if args.column_width is not None and args.column_width <= 0:
        print("--column-width must be positive", file=sys.stderr)
        return 2
    if args.due_date:
        try:
            datetime.strptime(args.due_date, "%Y-%m-%d")
        except ValueError:
            print("--due-date must use YYYY-MM-DD", file=sys.stderr)
            return 2
    try:
        kinds = parse_assignments(args.figure_kind, set(DEFAULT_THRESHOLDS))
        raw_widths = parse_assignments(args.figure_width)
        widths = {key: float(value) for key, value in raw_widths.items()}
        if any(value <= 0 for value in widths.values()):
            raise ValueError("Figure widths must be positive")
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 2

    args.output.parent.mkdir(parents=True, exist_ok=True)
    temp_root = args.output.parent / f".{args.output.name}.tmp-{uuid.uuid4().hex[:8]}"
    temp_root.mkdir()
    ctx = Context(args=args, root=temp_root, upload=temp_root / "upload", qa=temp_root / "qa")
    ctx.upload.mkdir()
    ctx.qa.mkdir()
    try:
        suffix = args.source.suffix.lower()
        if suffix == ".docx":
            process_docx(args.source, ctx, kinds, widths)
        elif suffix == ".zip":
            process_latex_zip(args.source, ctx, kinds, widths)
        else:
            ctx.error("UNSUPPORTED_SOURCE", "Source must be a .docx or Overleaf .zip")
        exit_code = finalize(ctx)
        temp_root.replace(args.output)
        print(f"{('BLOCKED' if exit_code else 'PASS')}: {args.output}")
        return exit_code
    except Exception as exc:
        ctx.error("UNEXPECTED_FAILURE", f"{type(exc).__name__}: {exc}")
        write_report(ctx)
        write_manifest(ctx, None)
        temp_root.replace(args.output)
        print(f"BLOCKED: {args.output} ({exc})", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
