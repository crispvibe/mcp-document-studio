import csv
import io
import json
import os
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from abc import ABC, abstractmethod
from collections.abc import Iterator
from html.parser import HTMLParser
from pathlib import Path, PurePosixPath
from typing import TYPE_CHECKING, cast
from urllib.parse import unquote
from xml.etree import ElementTree

from docx import Document as DocxDocument
from docx.document import Document as DocxDocumentType
from docx.oxml.ns import qn
from docx.parts.image import ImagePart
from docx.table import Table as DocxTable
from docx.text.paragraph import Paragraph as DocxParagraph
from mcp.server.fastmcp import FastMCP
from openpyxl import Workbook, load_workbook
from pypdf import PdfReader as PyPdfReader
from typing_extensions import override

if TYPE_CHECKING:
    from pptx.shapes.placeholder import SlidePlaceholder

mcp = FastMCP("Document Reader")

TEXT_ENCODINGS = ("utf-8", "utf-8-sig", "gbk", "gb2312", "latin-1")
PPTX_SLIDE_PATTERN = re.compile(r"ppt/slides/slide(\d+)\.xml$")
LIBREOFFICE_EXPORT_FILTERS = {
    ".doc": "MS Word 97",
    ".ppt": "MS PowerPoint 97",
}


def _format_docx_content_with_images(
    text_content: str, image_payload: dict[str, object]
) -> str:
    images = image_payload.get("images", [])
    if not isinstance(images, list) or not images:
        return text_content

    image_lines = ["", "=== Extracted Images ==="]
    output_dir = image_payload.get("output_dir")
    if isinstance(output_dir, str) and output_dir:
        image_lines.append(f"Temporary image directory: {output_dir}")

    for image in images:
        if not isinstance(image, dict):
            continue
        index = image.get("index")
        filename = image.get("saved_filename") or image.get("source_name") or "unknown"
        saved_path = image.get("saved_path") or ""
        saved_uri = image.get("saved_uri") or ""
        width_px = image.get("width_px")
        height_px = image.get("height_px")
        image_lines.append(
            f"[{index}] {filename} | {width_px}x{height_px}"
            f" | path: {saved_path} | uri: {saved_uri}"
        )

    if text_content:
        return text_content + "\n" + "\n".join(image_lines)
    return "\n".join(image_lines).strip()


def _read_text_file(file_path: str, *, empty_message: str, error_prefix: str) -> str:
    for encoding in TEXT_ENCODINGS:
        try:
            with open(file_path, "r", encoding=encoding, newline="") as file:
                text = file.read()
            # a BOM survives plain utf-8 decoding; drop it before returning
            if text.startswith("\ufeff"):
                text = text[1:]
            return text if text else empty_message
        except UnicodeDecodeError:
            continue
        except Exception as exc:
            return f"{error_prefix}: {str(exc)}"

    return f"{error_prefix}: Could not decode file with any supported encoding."


def _normalize_text_chunks(chunks: list[str], empty_message: str) -> str:
    normalized = [chunk.strip() for chunk in chunks if chunk and chunk.strip()]
    if not normalized:
        return empty_message
    return "\n".join(normalized)


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _extract_markup_text(content: bytes) -> str:
    root = ElementTree.fromstring(content)
    chunks = [" ".join(text.split()) for text in root.itertext()]
    return _normalize_text_chunks(chunks, "")


def _run_text_command(command: list[str]) -> str | None:
    try:
        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            check=False,
            timeout=30,
        )
    except Exception:
        return None

    if completed.returncode != 0:
        return None

    output = completed.stdout.strip()
    if not output or output == "(null)":
        return None

    return output


def _extract_text_with_mdls(file_path: str) -> str | None:
    if sys.platform != "darwin":
        return None

    executable = shutil.which("mdls")
    if executable is None:
        return None

    return _run_text_command(
        [executable, "-name", "kMDItemTextContent", "-raw", file_path]
    )


def _extract_text_with_textutil(file_path: str) -> str | None:
    executable = shutil.which("textutil")
    if executable is None:
        return None

    return _run_text_command([executable, "-convert", "txt", "-stdout", file_path])


def _extract_text_with_command(file_path: str, command_name: str) -> str | None:
    executable = shutil.which(command_name)
    if executable is None:
        return None

    return _run_text_command([executable, file_path])


def _common_text_run_chars() -> frozenset[str]:
    ascii_set = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    punct = " \t'\".,;:!?()[]{}<>+-*/=@#$%^&_|~\\`\n"
    cjk_punct = "，。；：！？（）【】《》、“”‘’—…·￥"
    return frozenset(ascii_set + punct + cjk_punct)


_COMMON_TEXT_CHARS = _common_text_run_chars()


def _is_common_text_char(char: str) -> bool:
    if char in _COMMON_TEXT_CHARS:
        return True
    code = ord(char)
    # CJK Unified Ideographs, CJK Symbols, Hiragana/Katakana, Hangul
    return (
        0x3000 <= code <= 0x9FFF
        or 0xAC00 <= code <= 0xD7AF
        or 0xFF00 <= code <= 0xFFEF
        or 0x2000 <= code <= 0x206F
    )


def _printable_text_runs(text: str, min_length: int = 4) -> list[str]:
    runs: list[str] = []
    for chunk in re.split(r"[\x00-\x08\x0b\x0c\x0e-\x1f]+", text):
        cleaned = re.sub(r"[^\S\n]+", " ", chunk).strip()
        if len(cleaned) < min_length:
            continue
        common = sum(1 for char in cleaned if _is_common_text_char(char))
        if common / len(cleaned) >= 0.7:
            runs.append(cleaned)
    return runs


def _extract_legacy_binary_text(file_path: str) -> str | None:
    """Best-effort text extraction from legacy OLE binary formats (.doc/.ppt).

    Legacy Office binaries store text as UTF-16LE or single-byte runs inside
    an OLE compound file. Pulling printable runs is a heuristic fallback used
    only when no dedicated extractor (antiword, catppt, LibreOffice) exists.
    """
    try:
        data = Path(file_path).read_bytes()
    except Exception:
        return None
    if not data:
        return None

    runs: list[str] = []
    seen: set[str] = set()
    for decoded in (
        data.decode("utf-16-le", errors="ignore"),
        data.decode("latin-1", errors="ignore"),
    ):
        for run in _printable_text_runs(decoded):
            key = run.lower()
            if run not in runs and key not in seen:
                seen.add(key)
                runs.append(run)

    if not runs:
        return None
    return "\n".join(runs[:500])


def _extract_text_with_libreoffice(file_path: str) -> str | None:
    executable = shutil.which("soffice") or shutil.which("libreoffice")
    if executable is None:
        return None

    with tempfile.TemporaryDirectory(prefix="mcp-document-reader-convert-") as temp_dir:
        completed = subprocess.run(
            [
                executable,
                "--headless",
                "--convert-to",
                "txt:Text",
                "--outdir",
                temp_dir,
                file_path,
            ],
            capture_output=True,
            text=True,
            check=False,
            timeout=60,
        )
        if completed.returncode != 0:
            return None

        converted_path = Path(temp_dir) / f"{Path(file_path).stem}.txt"
        if not converted_path.exists():
            return None

        content = _read_text_file(
            str(converted_path),
            empty_message="",
            error_prefix="Error reading converted document",
        )
        if not content or content.startswith("Error reading converted document:"):
            return None
        return content


def _extract_epub_document_paths(archive: zipfile.ZipFile) -> list[str]:
    container_xml = archive.read("META-INF/container.xml")
    container_root = ElementTree.fromstring(container_xml)
    namespace = {"container": "urn:oasis:names:tc:opendocument:xmlns:container"}
    rootfile = container_root.find(".//container:rootfile", namespace)
    if rootfile is None:
        raise ValueError("EPUB rootfile not found")

    opf_path = rootfile.attrib.get("full-path")
    if not opf_path:
        raise ValueError("EPUB package path is missing")

    opf_root = ElementTree.fromstring(archive.read(opf_path))
    opf_namespace = {"opf": "http://www.idpf.org/2007/opf"}
    opf_parent = PurePosixPath(opf_path).parent

    manifest: dict[str, str] = {}
    for item in opf_root.findall(".//opf:manifest/opf:item", opf_namespace):
        item_id = item.attrib.get("id")
        href = item.attrib.get("href")
        if not item_id or not href:
            continue
        href_path = unquote(href.split("#", 1)[0])
        manifest[item_id] = posixpath.normpath(str(opf_parent / href_path))

    document_paths: list[str] = []
    for itemref in opf_root.findall(".//opf:spine/opf:itemref", opf_namespace):
        item_id = itemref.attrib.get("idref")
        if item_id and item_id in manifest:
            document_paths.append(manifest[item_id])

    if document_paths:
        return document_paths

    fallback_paths = []
    for name in archive.namelist():
        lower_name = name.lower()
        if lower_name.endswith((".xhtml", ".html", ".htm")):
            fallback_paths.append(name)
    return sorted(fallback_paths)


def _ensure_parent_directory(file_path: Path) -> None:
    file_path.parent.mkdir(parents=True, exist_ok=True)


def _stringify_value(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, (int, float, bool)):
        return str(value)
    return json.dumps(value, ensure_ascii=False)


def _coerce_text_list(values: object) -> list[str]:
    if not isinstance(values, list):
        return []

    normalized: list[str] = []
    for value in values:
        text = _stringify_value(value).strip()
        if text:
            normalized.append(text)
    return normalized


def _coerce_table_rows(values: object) -> list[list[str]]:
    if not isinstance(values, list):
        return []

    normalized_rows: list[list[str]] = []
    for row in values:
        if not isinstance(row, list):
            continue
        normalized_row = [_stringify_value(cell) for cell in row]
        if any(cell.strip() for cell in normalized_row):
            normalized_rows.append(normalized_row)
    return normalized_rows


def _convert_with_libreoffice(
    source_path: Path, target_path: Path, filter_name: str | None = None
) -> Path | None:
    executable = shutil.which("soffice") or shutil.which("libreoffice")
    if executable is None:
        return None

    _ensure_parent_directory(target_path)
    convert_to = target_path.suffix.lstrip(".")
    if filter_name:
        convert_to = f"{convert_to}:{filter_name}"

    completed = subprocess.run(
        [
            executable,
            "--headless",
            "--convert-to",
            convert_to,
            "--outdir",
            str(target_path.parent),
            str(source_path),
        ],
        capture_output=True,
        text=True,
        check=False,
        timeout=120,
    )
    if completed.returncode != 0:
        return None

    if not target_path.exists():
        return None
    return target_path


def _build_word_document(
    title: str | None,
    paragraphs: list[str] | None,
    tables: list[dict[str, object]] | None,
) -> DocxDocumentType:
    document = DocxDocument()

    if title and title.strip():
        document.add_heading(title.strip(), level=0)

    for paragraph in paragraphs or []:
        text = _stringify_value(paragraph).strip()
        if text:
            document.add_paragraph(text)

    for table_spec in tables or []:
        table_title = _stringify_value(table_spec.get("title")).strip()
        if table_title:
            document.add_paragraph(table_title)

        headers = _coerce_text_list(table_spec.get("headers"))
        rows = _coerce_table_rows(table_spec.get("rows"))
        column_count = max(len(headers), max((len(row) for row in rows), default=0))
        if column_count == 0:
            continue

        total_rows = (1 if headers else 0) + len(rows)
        table = document.add_table(rows=total_rows, cols=column_count)
        table.style = "Table Grid"

        if headers:
            for column_index in range(column_count):
                table.rows[0].cells[column_index].text = (
                    headers[column_index] if column_index < len(headers) else ""
                )

        row_offset = 1 if headers else 0
        for row_index, row in enumerate(rows, start=row_offset):
            for column_index in range(column_count):
                table.rows[row_index].cells[column_index].text = (
                    row[column_index] if column_index < len(row) else ""
                )

    return document


def _load_presentation_dependencies():
    try:
        from pptx import Presentation
        from pptx.util import Inches
    except ImportError as exc:
        raise RuntimeError(
            "python-pptx is required for presentation generation. "
            "Install project dependencies first."
        ) from exc

    return Presentation, Inches


def _build_presentation(
    title: str | None,
    subtitle: str | None,
    slides: list[dict[str, object]] | None,
):
    Presentation, Inches = _load_presentation_dependencies()
    presentation = Presentation()

    if title or subtitle:
        title_slide = presentation.slides.add_slide(presentation.slide_layouts[0])
        title_shape = title_slide.shapes.title
        if title_shape is not None and title_shape.has_text_frame:
            title_shape.text_frame.text = title or ""
        if len(title_slide.placeholders) > 1:
            subtitle_shape = cast("SlidePlaceholder", title_slide.placeholders[1])
            if subtitle_shape.has_text_frame:
                subtitle_shape.text_frame.text = subtitle or ""

    for slide_spec in slides or []:
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        current_top = 0.4

        slide_title = _stringify_value(slide_spec.get("title")).strip()
        if slide_title:
            title_box = slide.shapes.add_textbox(
                Inches(0.5), Inches(current_top), Inches(8.5), Inches(0.6)
            )
            title_box.text_frame.text = slide_title
            current_top += 0.8

        paragraphs = _coerce_text_list(slide_spec.get("paragraphs"))
        bullets = _coerce_text_list(slide_spec.get("bullets"))
        table_spec = slide_spec.get("table")

        if paragraphs or bullets:
            text_height = 2.8 if isinstance(table_spec, dict) else 4.8
            text_box = slide.shapes.add_textbox(
                Inches(0.6), Inches(current_top), Inches(8.2), Inches(text_height)
            )
            text_frame = text_box.text_frame
            first_paragraph = True

            for paragraph in paragraphs:
                text_paragraph = (
                    text_frame.paragraphs[0]
                    if first_paragraph
                    else text_frame.add_paragraph()
                )
                text_paragraph.text = paragraph
                first_paragraph = False

            for bullet in bullets:
                text_paragraph = (
                    text_frame.paragraphs[0]
                    if first_paragraph
                    else text_frame.add_paragraph()
                )
                text_paragraph.text = bullet
                text_paragraph.level = 0
                first_paragraph = False

            current_top += text_height + 0.2

        if isinstance(table_spec, dict):
            headers = _coerce_text_list(table_spec.get("headers"))
            rows = _coerce_table_rows(table_spec.get("rows"))
            column_count = max(len(headers), max((len(row) for row in rows), default=0))
            total_rows = (1 if headers else 0) + len(rows)
            if column_count > 0 and total_rows > 0:
                table_shape = slide.shapes.add_table(
                    total_rows,
                    column_count,
                    Inches(0.6),
                    Inches(current_top),
                    Inches(8.2),
                    Inches(max(1.2, min(4.5, total_rows * 0.45))),
                )
                table = table_shape.table

                if headers:
                    for column_index in range(column_count):
                        table.cell(0, column_index).text = (
                            headers[column_index] if column_index < len(headers) else ""
                        )

                row_offset = 1 if headers else 0
                for row_index, row in enumerate(rows, start=row_offset):
                    for column_index in range(column_count):
                        table.cell(row_index, column_index).text = (
                            row[column_index] if column_index < len(row) else ""
                        )

    if not presentation.slides:
        presentation.slides.add_slide(presentation.slide_layouts[6])

    return presentation


_INVALID_SHEET_NAME_CHARS = re.compile(r"[\\/*?:\[\]]")


def _sanitize_sheet_name(name: object, default: str) -> str:
    cleaned = _INVALID_SHEET_NAME_CHARS.sub("", _stringify_value(name)).strip()
    return cleaned[:31] or default


def _coerce_cell_value(value: object) -> str | int | float | bool | None:
    if value is None:
        return None
    if isinstance(value, (str, int, float, bool)):
        return value
    return _stringify_value(value)


def _coerce_sheet_rows(values: object) -> list[list[str | int | float | bool | None]]:
    if not isinstance(values, list):
        return []
    return [
        [_coerce_cell_value(cell) for cell in row]
        for row in values
        if isinstance(row, list)
    ]


def _build_workbook(sheets: list[dict[str, object]] | None) -> Workbook:
    workbook = Workbook()
    first_sheet = True

    for index, sheet_spec in enumerate(sheets or [], start=1):
        if not isinstance(sheet_spec, dict):
            continue
        name = _sanitize_sheet_name(sheet_spec.get("name"), f"Sheet{index}")
        if first_sheet:
            worksheet = workbook.active
            if worksheet is None:
                worksheet = workbook.create_sheet()
        else:
            worksheet = workbook.create_sheet()
        worksheet.title = name

        headers = _coerce_text_list(sheet_spec.get("headers"))
        if headers:
            worksheet.append(headers)
        for row in _coerce_sheet_rows(sheet_spec.get("rows")):
            worksheet.append(row)
        first_sheet = False

    return workbook


def _write_csv_file(
    target_path: Path,
    headers: list[str],
    rows: list[list[str | int | float | bool | None]],
) -> None:
    # utf-8-sig so spreadsheet apps detect UTF-8 correctly
    with open(target_path, "w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file)
        if headers:
            writer.writerow(headers)
        for row in rows:
            writer.writerow(["" if cell is None else cell for cell in row])


def _serialize_generated_file(
    file_path: Path, generated_format: str, source_format: str | None = None
) -> str:
    payload: dict[str, object] = {
        "path": str(file_path.resolve()),
        "uri": file_path.resolve().as_uri(),
        "format": generated_format,
    }
    if source_format is not None:
        payload["source_format"] = source_format
    return json.dumps(payload, ensure_ascii=False, indent=2)


class DocumentReader(ABC):
    """Abstract base class for document readers"""

    @abstractmethod
    def read(self, file_path: str) -> str:
        """Read and extract text from a document"""
        pass


def _iter_docx_block_items(
    document: DocxDocumentType,
) -> Iterator[DocxParagraph | DocxTable]:
    """Yield paragraphs and tables in document order."""
    body = document.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield DocxParagraph(child, document)
        elif child.tag == qn("w:tbl"):
            yield DocxTable(child, document)


def _docx_table_rows_text(table: DocxTable) -> list[str]:
    rows_text: list[str] = []
    for row in table.rows:
        row_text = []
        for cell in row.cells:
            cell_text = " ".join(p.text for p in cell.paragraphs).strip()
            if cell_text:
                row_text.append(cell_text)
        if row_text:
            rows_text.append("\t".join(row_text))
    return rows_text


class DocxReader(DocumentReader):
    """DOCX document reader implementation"""

    @override
    def read(self, file_path: str) -> str:
        """Read and extract text from DOCX file"""
        try:
            doc = DocxDocument(file_path)
            text = []

            for block in _iter_docx_block_items(doc):
                if isinstance(block, DocxParagraph):
                    if block.text:
                        text.append(block.text)
                else:
                    text.extend(_docx_table_rows_text(block))

            extracted_text = "\n".join(text)
            return extracted_text if extracted_text else "No text found in the DOCX."
        except Exception as e:
            return f"Error reading DOCX: {str(e)}"

    def extract_images(
        self, file_path: str, output_dir: str | None = None
    ) -> dict[str, object]:
        doc = DocxDocument(file_path)
        file_path_obj = Path(file_path).resolve()
        extracted_images: list[dict[str, object]] = []

        for index in range(len(doc.inline_shapes)):
            shape = doc.inline_shapes[index]
            pic = getattr(shape._inline.graphic.graphicData, "pic", None)
            if pic is None or getattr(pic.blipFill, "blip", None) is None:
                continue

            rel_id = pic.blipFill.blip.embed
            if not rel_id or rel_id not in doc.part.rels:
                continue

            rel = doc.part.rels[rel_id]
            if "image" not in rel.reltype or rel.is_external:
                continue

            image_part = rel.target_part
            if not isinstance(image_part, ImagePart):
                continue
            image = image_part.image
            docpr = getattr(shape._inline, "docPr", None)
            extracted_images.append(
                {
                    "index": index + 1,
                    "rel_id": rel_id,
                    "source_name": Path(str(image_part.partname)).name,
                    "content_type": image.content_type,
                    "width_px": image.px_width,
                    "height_px": image.px_height,
                    "width_emu": int(shape.width),
                    "height_emu": int(shape.height),
                    "size_bytes": len(image_part.blob),
                    "shape_name": getattr(docpr, "name", None),
                    "title": getattr(docpr, "title", None),
                    "description": getattr(docpr, "descr", None),
                    "blob": image_part.blob,
                }
            )

        export_dir: Path | None = None
        if output_dir is not None:
            export_dir = Path(output_dir).expanduser().resolve()
            export_dir.mkdir(parents=True, exist_ok=True)
        elif extracted_images:
            export_dir = Path(
                tempfile.mkdtemp(prefix="mcp-document-reader-images-")
            ).resolve()

        images: list[dict[str, object]] = []
        for position, image_info in enumerate(extracted_images, start=1):
            blob = cast(bytes, image_info.pop("blob"))
            saved_path: str | None = None
            saved_uri: str | None = None
            saved_filename: str | None = None

            if export_dir is not None:
                source_name = str(image_info["source_name"])
                saved_filename = f"{position:03d}_{source_name}"
                output_path = export_dir / saved_filename
                output_path.write_bytes(blob)
                resolved_output_path = output_path.resolve()
                saved_path = str(resolved_output_path)
                saved_uri = resolved_output_path.as_uri()

            images.append(
                {
                    **image_info,
                    "saved_filename": saved_filename,
                    "saved_path": saved_path,
                    "saved_uri": saved_uri,
                }
            )

        return {
            "document": str(file_path_obj),
            "document_uri": file_path_obj.as_uri(),
            "output_dir": str(export_dir) if export_dir is not None else None,
            "output_dir_uri": export_dir.as_uri() if export_dir is not None else None,
            "image_count": len(images),
            "images": images,
        }


class PdfReader(DocumentReader):
    """PDF document reader implementation"""

    @override
    def read(self, file_path: str) -> str:
        """Read and extract text from PDF file"""
        try:
            with open(file_path, "rb") as file:
                pdf_reader = PyPdfReader(file)
                if pdf_reader.is_encrypted:
                    try:
                        pdf_reader.decrypt("")
                    except Exception:
                        pass
                text = []

                for page in pdf_reader.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text.append(page_text.strip())

                extracted_text = "\n\n".join(text)
                if extracted_text:
                    return extracted_text
                return (
                    "No text found in the PDF. "
                    "Scanned or image-only PDFs require OCR, which is not bundled."
                )
        except Exception as e:
            return f"Error reading PDF: {str(e)}"


class TxtReader(DocumentReader):
    """TXT document reader implementation"""

    @override
    def read(self, file_path: str) -> str:
        """Read and extract text from TXT file with encoding handling"""
        return _read_text_file(
            file_path,
            empty_message="No text found in the TXT file.",
            error_prefix="Error reading TXT",
        )


class CsvReader(DocumentReader):
    @override
    def read(self, file_path: str) -> str:
        for encoding in TEXT_ENCODINGS:
            try:
                with open(file_path, "r", encoding=encoding, newline="") as file:
                    content = file.read()
                rows = _csv_text_to_rows(content)
                return _normalize_text_chunks(rows, "No text found in the CSV file.")
            except UnicodeDecodeError:
                continue
            except Exception as exc:
                return f"Error reading CSV: {str(exc)}"

        return "Error reading CSV: Could not decode file with any supported encoding."


class MarkdownReader(DocumentReader):
    @override
    def read(self, file_path: str) -> str:
        return _read_text_file(
            file_path,
            empty_message="No text found in the Markdown file.",
            error_prefix="Error reading Markdown",
        )


class DocReader(DocumentReader):
    @override
    def read(self, file_path: str) -> str:
        try:
            extracted = (
                _extract_text_with_mdls(file_path)
                or _extract_text_with_textutil(file_path)
                or _extract_text_with_command(file_path, "antiword")
                or _extract_text_with_libreoffice(file_path)
                or _extract_legacy_binary_text(file_path)
            )
            if extracted:
                return extracted
            return (
                "Error reading DOC: No text could be extracted. "
                "Try installing antiword or LibreOffice for better results."
            )
        except Exception as exc:
            return f"Error reading DOC: {str(exc)}"


class PptReader(DocumentReader):
    @override
    def read(self, file_path: str) -> str:
        try:
            extracted = (
                _extract_text_with_mdls(file_path)
                or _extract_text_with_command(file_path, "catppt")
                or _extract_text_with_libreoffice(file_path)
                or _extract_legacy_binary_text(file_path)
            )
            if extracted:
                return extracted
            return (
                "Error reading PPT: No text could be extracted. "
                "Try installing catdoc/catppt or LibreOffice for better results."
            )
        except Exception as exc:
            return f"Error reading PPT: {str(exc)}"


def _get_pptx_slide_paths(archive: zipfile.ZipFile) -> list[str]:
    """Return slide part paths in presentation order.

    Resolves the p:sldIdLst ordering from ppt/presentation.xml through
    presentation.xml.rels. Falls back to numeric filename ordering when the
    manifest is missing or incomplete.
    """
    names = set(archive.namelist())
    try:
        presentation_root = ElementTree.fromstring(archive.read("ppt/presentation.xml"))
        rels_root = ElementTree.fromstring(
            archive.read("ppt/_rels/presentation.xml.rels")
        )
    except KeyError:
        presentation_root = None
        rels_root = None

    if presentation_root is not None and rels_root is not None:
        rel_targets: dict[str, str] = {}
        for rel in rels_root:
            rel_id = rel.attrib.get("Id")
            target = rel.attrib.get("Target")
            if not rel_id or not target:
                continue
            if target.startswith("/"):
                rel_targets[rel_id] = target.lstrip("/")
            else:
                rel_targets[rel_id] = posixpath.normpath(f"ppt/{target}")

        rel_id_key = (
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        )
        ordered: list[str] = []
        for slide_id in presentation_root.iter():
            if _local_name(slide_id.tag) != "sldId":
                continue
            rel_id = slide_id.attrib.get(rel_id_key)
            slide_path = rel_targets.get(rel_id) if rel_id else None
            if slide_path and slide_path in names:
                ordered.append(slide_path)
        if ordered:
            return ordered

    numbered = [
        (match, name)
        for name in names
        if (match := PPTX_SLIDE_PATTERN.match(name)) is not None
    ]
    return [
        name for match, name in sorted(numbered, key=lambda item: int(item[0].group(1)))
    ]


class PptxReader(DocumentReader):
    @override
    def read(self, file_path: str) -> str:
        try:
            with zipfile.ZipFile(file_path) as archive:
                ordered_slide_names = _get_pptx_slide_paths(archive)
                text: list[str] = []

                for index, slide_name in enumerate(ordered_slide_names, start=1):
                    slide_text = _extract_markup_text(archive.read(slide_name))
                    if not slide_text:
                        continue
                    text.append(f"=== Slide: {index} ===")
                    text.append(slide_text)
                    text.append("")

                return _normalize_text_chunks(text, "No text found in the PPTX.")
        except Exception as exc:
            return f"Error reading PPTX: {str(exc)}"


class EpubReader(DocumentReader):
    @override
    def read(self, file_path: str) -> str:
        try:
            with zipfile.ZipFile(file_path) as archive:
                document_paths = _extract_epub_document_paths(archive)
                text: list[str] = []

                for index, document_path in enumerate(document_paths, start=1):
                    document_text = _extract_markup_text(archive.read(document_path))
                    if not document_text:
                        continue
                    text.append(f"=== Section: {index} ===")
                    text.append(document_text)
                    text.append("")

                return _normalize_text_chunks(text, "No text found in the EPUB.")
        except Exception as exc:
            return f"Error reading EPUB: {str(exc)}"


def _csv_text_to_rows(text: str) -> list[str]:
    rows: list[str] = []
    for row in csv.reader(io.StringIO(text.lstrip("\ufeff"))):
        row_text = [cell.strip() for cell in row]
        if any(row_text):
            rows.append("\t".join(row_text))
    return rows


class ExcelReader(DocumentReader):
    """Excel document reader implementation"""

    _CALAMINE_SUFFIXES = {".xls", ".xlsb", ".ods"}

    @override
    def read(self, file_path: str) -> str:
        """Read and extract text from Excel file"""
        if Path(file_path).suffix.lower() in self._CALAMINE_SUFFIXES:
            return self._read_legacy_spreadsheet(file_path)
        return self._read_xlsx(file_path)

    def _read_legacy_spreadsheet(self, file_path: str) -> str:
        try:
            extracted = self._read_with_calamine(file_path)
            if extracted is not None:
                return extracted

            csv_text = _extract_text_with_command(file_path, "xls2csv")
            if csv_text:
                normalized = _normalize_text_chunks(_csv_text_to_rows(csv_text), "")
                if normalized:
                    return normalized

            converted = self._convert_xls_to_xlsx(file_path)
            if converted is not None:
                return converted

            binary_text = _extract_legacy_binary_text(file_path)
            if binary_text:
                return binary_text
            return (
                "Error reading Excel: could not extract content. "
                "Install LibreOffice or xls2csv for better results."
            )
        except Exception as exc:
            return f"Error reading Excel: {str(exc)}"

    def _read_with_calamine(self, file_path: str) -> str | None:
        try:
            from python_calamine import CalamineWorkbook
        except ImportError:
            return None

        try:
            workbook = CalamineWorkbook.from_path(file_path)
        except Exception:
            return None

        text: list[str] = []
        for sheet_name in workbook.sheet_names:
            sheet = workbook.get_sheet_by_name(sheet_name)
            text.append(f"=== Sheet: {sheet_name} ===")
            for row in sheet.to_python():
                row_text = [str(cell) if cell is not None else "" for cell in row]
                if any(cell.strip() for cell in row_text):
                    text.append("\t".join(row_text))
            text.append("")

        extracted = "\n".join(text)
        return extracted or None

    def _read_xlsx(self, file_path: str) -> str:
        try:
            wb = load_workbook(file_path, read_only=True)
            text = []

            for sheet_name in wb.sheetnames:
                sheet = wb[sheet_name]
                text.append(f"=== Sheet: {sheet_name} ===")

                for row in sheet.iter_rows(values_only=True):
                    row_text = [str(cell) if cell is not None else "" for cell in row]
                    if any(row_text):
                        text.append("\t".join(row_text))

                text.append("")

            extracted_text = "\n".join(text)
            wb.close()
            return (
                extracted_text if extracted_text else "No text found in the Excel file."
            )
        except Exception as e:
            return f"Error reading Excel: {str(e)}"

    def _convert_xls_to_xlsx(self, file_path: str) -> str | None:
        with tempfile.TemporaryDirectory(prefix="mcp-document-reader-xls-") as temp_dir:
            target_path = Path(temp_dir) / f"{Path(file_path).stem}.xlsx"
            converted_path = _convert_with_libreoffice(
                Path(file_path), target_path, "Calc MS Excel 2007 XML"
            )
            if converted_path is None:
                return None
            result = self._read_xlsx(str(converted_path))
            if result.startswith("Error reading Excel:"):
                return None
            return result


class _HtmlTextExtractor(HTMLParser):
    """Collect visible text while skipping script/style content."""

    _SKIP_TAGS = {"script", "style", "noscript", "template", "head"}
    _BLOCK_TAGS = {
        "address",
        "article",
        "aside",
        "blockquote",
        "br",
        "dd",
        "details",
        "div",
        "dl",
        "dt",
        "figcaption",
        "figure",
        "footer",
        "form",
        "h1",
        "h2",
        "h3",
        "h4",
        "h5",
        "h6",
        "header",
        "hr",
        "li",
        "main",
        "nav",
        "ol",
        "p",
        "pre",
        "section",
        "summary",
        "table",
        "td",
        "th",
        "tr",
        "ul",
    }

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self._chunks: list[str] = []
        self._skip_depth = 0

    @override
    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in self._SKIP_TAGS:
            self._skip_depth += 1
        elif tag in self._BLOCK_TAGS and self._skip_depth == 0:
            self._chunks.append("\n")

    @override
    def handle_endtag(self, tag: str) -> None:
        if tag in self._SKIP_TAGS and self._skip_depth > 0:
            self._skip_depth -= 1
        elif tag in self._BLOCK_TAGS and self._skip_depth == 0:
            self._chunks.append("\n")

    @override
    def handle_data(self, data: str) -> None:
        if self._skip_depth == 0:
            self._chunks.append(data)

    def text(self) -> str:
        return "".join(self._chunks)


class HtmlReader(DocumentReader):
    """HTML document reader implementation"""

    @override
    def read(self, file_path: str) -> str:
        for encoding in TEXT_ENCODINGS:
            try:
                with open(file_path, "r", encoding=encoding) as file:
                    markup = file.read()
            except UnicodeDecodeError:
                continue
            except Exception as exc:
                return f"Error reading HTML: {str(exc)}"

            extractor = _HtmlTextExtractor()
            try:
                extractor.feed(markup)
            except Exception:
                pass
            return _normalize_text_chunks(
                extractor.text().splitlines(), "No text found in the HTML file."
            )

        return "Error reading HTML: Could not decode file with any supported encoding."


class JsonReader(DocumentReader):
    """JSON document reader implementation"""

    @override
    def read(self, file_path: str) -> str:
        content = _read_text_file(
            file_path,
            empty_message="No text found in the JSON file.",
            error_prefix="Error reading JSON",
        )
        if content.startswith("Error reading JSON") or not content.strip():
            return content

        try:
            parsed = json.loads(content)
        except json.JSONDecodeError:
            return content
        return json.dumps(parsed, ensure_ascii=False, indent=2)


_RTF_SKIP_DESTINATIONS = {
    "fonttbl",
    "colortbl",
    "stylesheet",
    "info",
    "pict",
    "object",
    "header",
    "footer",
    "footnote",
    "annotation",
    "xmlnstbl",
    "listtable",
    "listoverridetable",
    "revtbl",
    "rsidtbl",
    "generator",
    "datastore",
    "themedata",
    "colorschememapping",
    "latentstyles",
    "filetbl",
}


def _rtf_to_text(markup: str) -> str:
    """Convert RTF markup to plain text (covers common RTF documents)."""
    chunks: list[str] = []
    stack: list[bool] = []  # per open group: is it an ignorable destination
    skipped_groups = 0
    index = 0
    length = len(markup)
    uc_skip = 1  # number of fallback chars to skip after \u control words
    pending_uc_skip = 0

    def _skip_active() -> bool:
        return skipped_groups > 0

    def _mark_current_group_skipped() -> None:
        nonlocal skipped_groups
        if stack and not stack[-1]:
            stack[-1] = True
            skipped_groups += 1

    while index < length:
        char = markup[index]
        if char == "{":
            stack.append(False)
            index += 1
            continue
        if char == "}":
            if stack and stack.pop():
                skipped_groups -= 1
            index += 1
            continue
        if char != "\\":
            if pending_uc_skip > 0:
                pending_uc_skip -= 1
            elif not _skip_active():
                chunks.append(char)
            index += 1
            continue

        # control symbol or control word
        index += 1
        if index >= length:
            break
        next_char = markup[index]
        if next_char in "{}\\":
            if not _skip_active():
                chunks.append(next_char)
            index += 1
            continue
        if next_char == "~":
            if not _skip_active():
                chunks.append(" ")
            index += 1
            continue
        if next_char == "*":
            # ignorable destination — mark current group as skipped
            _mark_current_group_skipped()
            index += 1
            continue
        if next_char == "'" and index + 2 < length:
            hex_digits = markup[index + 1 : index + 3]
            if not _skip_active():
                try:
                    chunks.append(bytes([int(hex_digits, 16)]).decode("cp1252"))
                except (ValueError, UnicodeDecodeError):
                    pass
            index += 3
            continue

        match = re.match(r"([a-zA-Z]+)(-?\d+)? ?", markup[index:])
        if match is None:
            index += 1
            continue
        word, argument = match.group(1), match.group(2)
        index += match.end()

        if word in _RTF_SKIP_DESTINATIONS:
            _mark_current_group_skipped()
            continue
        if word == "uc" and argument is not None:
            uc_skip = max(0, int(argument))
        if _skip_active():
            continue
        if word == "u" and argument is not None:
            code = int(argument)
            if code < 0:
                code += 65536
            try:
                chunks.append(chr(code))
            except ValueError:
                pass
            pending_uc_skip = uc_skip
        elif word in {"par", "line", "row", "page", "sect"}:
            chunks.append("\n")
        elif word == "tab":
            chunks.append("\t")
        elif word == "emdash":
            chunks.append("\u2014")
        elif word == "endash":
            chunks.append("\u2013")
        elif word in {"lquote", "rquote"}:
            chunks.append("'")
        elif word in {"ldblquote", "rdblquote"}:
            chunks.append('"')
        elif word == "bullet":
            chunks.append("\u2022")

    return "".join(chunks)


class RtfReader(DocumentReader):
    """RTF document reader implementation"""

    @override
    def read(self, file_path: str) -> str:
        content = _read_text_file(
            file_path,
            empty_message="No text found in the RTF file.",
            error_prefix="Error reading RTF",
        )
        if content.startswith("Error reading RTF") or not content.strip():
            return content
        return _normalize_text_chunks(
            _rtf_to_text(content).splitlines(), "No text found in the RTF file."
        )


class DocumentReaderFactory:
    """Factory for creating document readers based on file extension"""

    _readers: dict[str, type[DocumentReader]] = {
        ".txt": TxtReader,
        ".csv": CsvReader,
        ".md": MarkdownReader,
        ".markdown": MarkdownReader,
        ".doc": DocReader,
        ".docx": DocxReader,
        ".pdf": PdfReader,
        ".ppt": PptReader,
        ".pptx": PptxReader,
        ".epub": EpubReader,
        ".xlsx": ExcelReader,
        ".xls": ExcelReader,
        ".xlsb": ExcelReader,
        ".xlsm": ExcelReader,
        ".ods": ExcelReader,
        ".html": HtmlReader,
        ".htm": HtmlReader,
        ".json": JsonReader,
        ".xml": TxtReader,
        ".yaml": TxtReader,
        ".yml": TxtReader,
        ".rtf": RtfReader,
    }

    @classmethod
    def get_reader(cls, file_path: str) -> DocumentReader:
        """Get appropriate reader for the given file"""
        _, ext = os.path.splitext(file_path.lower())
        if ext not in cls._readers:
            raise ValueError(f"Unsupported document type: {ext}")
        return cls._readers[ext]()

    @classmethod
    def is_supported(cls, file_path: str) -> bool:
        """Check if the file type is supported"""
        _, ext = os.path.splitext(file_path.lower())
        return ext in cls._readers

    @classmethod
    def supported_extensions(cls) -> list[str]:
        """List all supported file extensions for reading."""
        return sorted(cls._readers)


WRITE_FORMATS: dict[str, list[str]] = {
    "word": [".docx", ".doc"],
    "presentation": [".pptx", ".ppt"],
    "spreadsheet": [".xlsx", ".csv", ".xls"],
}


@mcp.tool()
def read_document(filename: str) -> str:
    """
    Reads and extracts text from a specified document file.
    Supports TXT, CSV, Markdown, DOC, DOCX, PDF, PPT, PPTX, EPUB,
    and Excel (XLSX, XLS) files.

    :param filename: Path to the document file to read
        (supports absolute or relative paths)
    :return: Extracted text from the document
    """
    file_path = Path(filename).expanduser()

    if not file_path.exists():
        return f"Error: File '{filename}' not found."

    if not DocumentReaderFactory.is_supported(str(file_path)):
        return f"Error: Unsupported document type for file '{filename}'."

    try:
        reader = DocumentReaderFactory.get_reader(str(file_path))
        content = reader.read(str(file_path))
        if isinstance(reader, DocxReader) and not content.startswith(
            "Error reading DOCX:"
        ):
            image_payload = reader.extract_images(str(file_path))
            return _format_docx_content_with_images(content, image_payload)
        return content
    except Exception as e:
        return f"Error reading document: {str(e)}"


@mcp.tool()
def extract_document_images(filename: str, output_dir: str | None = None) -> str:
    """
    Extracts embedded images from a DOCX file and returns structured JSON metadata.

    :param filename: Path to the DOCX document
    :param output_dir: Optional directory to save extracted images
    :return: JSON payload containing extracted image metadata and saved file paths
    """
    file_path = Path(filename).expanduser()

    if not file_path.exists():
        return f"Error: File '{filename}' not found."

    if file_path.suffix.lower() != ".docx":
        return "Error: Image extraction currently supports DOCX files only."

    try:
        reader = DocxReader()
        result = reader.extract_images(str(file_path), output_dir)
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as e:
        return f"Error extracting document images: {str(e)}"


@mcp.tool()
def write_word_document(
    filename: str,
    title: str | None = None,
    paragraphs: list[str] | None = None,
    tables: list[dict[str, object]] | None = None,
) -> str:
    """
    Generates a Word document in DOCX format, or DOC via LibreOffice conversion.

    :param filename: Target output path ending with .docx or .doc
    :param title: Optional document title
    :param paragraphs: Optional paragraph list written in order
    :param tables: Optional table definitions using title, headers, and rows
    :return: JSON payload describing the generated file path and format
    """
    target_path = Path(filename).expanduser()
    target_suffix = target_path.suffix.lower()
    if target_suffix not in {".docx", ".doc"}:
        return "Error: Word generation supports .docx and .doc output only."

    try:
        _ensure_parent_directory(target_path)
        document = _build_word_document(title, paragraphs, tables)
        if target_suffix == ".docx":
            document.save(str(target_path))
            return _serialize_generated_file(target_path, "docx")

        with tempfile.TemporaryDirectory(
            prefix="mcp-document-writer-word-"
        ) as temp_dir:
            temp_docx_path = Path(temp_dir) / f"{target_path.stem}.docx"
            document.save(str(temp_docx_path))
            converted_path = _convert_with_libreoffice(
                temp_docx_path,
                target_path,
                LIBREOFFICE_EXPORT_FILTERS.get(".doc"),
            )

        if converted_path is None:
            return (
                "Error writing DOC: LibreOffice conversion is unavailable. "
                "Install LibreOffice and ensure soffice/libreoffice is in PATH."
            )
        return _serialize_generated_file(converted_path, "doc", source_format="docx")
    except Exception as exc:
        return f"Error writing Word document: {str(exc)}"


@mcp.tool()
def write_presentation(
    filename: str,
    title: str | None = None,
    subtitle: str | None = None,
    slides: list[dict[str, object]] | None = None,
) -> str:
    """
    Generates a PowerPoint presentation in PPTX format, or PPT via conversion.

    :param filename: Target output path ending with .pptx or .ppt
    :param title: Optional title slide title
    :param subtitle: Optional title slide subtitle
    :param slides: Optional slide definitions (title, paragraphs, bullets, table)
    :return: JSON payload describing the generated file path and format
    """
    target_path = Path(filename).expanduser()
    target_suffix = target_path.suffix.lower()
    if target_suffix not in {".pptx", ".ppt"}:
        return "Error: Presentation generation supports .pptx and .ppt output only."

    try:
        _ensure_parent_directory(target_path)
        presentation = _build_presentation(title, subtitle, slides)
        if target_suffix == ".pptx":
            presentation.save(str(target_path))
            return _serialize_generated_file(target_path, "pptx")

        with tempfile.TemporaryDirectory(prefix="mcp-document-writer-ppt-") as temp_dir:
            temp_pptx_path = Path(temp_dir) / f"{target_path.stem}.pptx"
            presentation.save(str(temp_pptx_path))
            converted_path = _convert_with_libreoffice(
                temp_pptx_path,
                target_path,
                LIBREOFFICE_EXPORT_FILTERS.get(".ppt"),
            )

        if converted_path is None:
            return (
                "Error writing PPT: LibreOffice conversion is unavailable. "
                "Install LibreOffice and ensure soffice/libreoffice is in PATH."
            )
        return _serialize_generated_file(converted_path, "ppt", source_format="pptx")
    except Exception as exc:
        return f"Error writing presentation: {str(exc)}"


@mcp.tool()
def write_spreadsheet(
    filename: str,
    sheets: list[dict[str, object]] | None = None,
    headers: list[str] | None = None,
    rows: list[list[object]] | None = None,
) -> str:
    """
    Generates a spreadsheet in XLSX or CSV format, or XLS via LibreOffice conversion.

    :param filename: Target output path ending with .xlsx, .csv, or .xls
    :param sheets: Optional sheet definitions with name, headers, and rows
    :param headers: Optional single-sheet headers (used when sheets is empty)
    :param rows: Optional single-sheet rows (used when sheets is empty)
    :return: JSON payload describing the generated file path and format
    """
    target_path = Path(filename).expanduser()
    target_suffix = target_path.suffix.lower()
    if target_suffix not in {".xlsx", ".csv", ".xls"}:
        return (
            "Error: Spreadsheet generation supports .xlsx, .csv and .xls output only."
        )

    sheet_specs: list[dict[str, object]]
    if sheets:
        sheet_specs = [spec for spec in sheets if isinstance(spec, dict)]
    else:
        sheet_specs = [{"name": "Sheet1", "headers": headers, "rows": rows}]

    if target_suffix == ".csv" and len(sheet_specs) > 1:
        return "Error: CSV output supports a single sheet only."

    try:
        _ensure_parent_directory(target_path)
        if target_suffix == ".csv":
            sheet = sheet_specs[0]
            _write_csv_file(
                target_path,
                _coerce_text_list(sheet.get("headers")),
                _coerce_sheet_rows(sheet.get("rows")),
            )
            return _serialize_generated_file(target_path, "csv")

        workbook = _build_workbook(sheet_specs)
        if target_suffix == ".xlsx":
            workbook.save(str(target_path))
            return _serialize_generated_file(target_path, "xlsx")

        with tempfile.TemporaryDirectory(prefix="mcp-document-writer-xls-") as temp_dir:
            temp_xlsx_path = Path(temp_dir) / f"{target_path.stem}.xlsx"
            workbook.save(str(temp_xlsx_path))
            converted_path = _convert_with_libreoffice(
                temp_xlsx_path,
                target_path,
                "MS Excel 97",
            )

        if converted_path is None:
            return (
                "Error writing XLS: LibreOffice conversion is unavailable. "
                "Install LibreOffice and ensure soffice/libreoffice is in PATH."
            )
        return _serialize_generated_file(converted_path, "xls", source_format="xlsx")
    except Exception as exc:
        return f"Error writing spreadsheet: {str(exc)}"


@mcp.tool()
def convert_document(
    filename: str,
    target_format: str,
    output_dir: str | None = None,
) -> str:
    """
    Converts a document to another format via LibreOffice.

    :param filename: Source document path
    :param target_format: Target extension such as pdf, docx, txt, html, csv
    :param output_dir: Optional output directory (defaults to source directory)
    :return: JSON payload describing the converted file path and format
    """
    source_path = Path(filename).expanduser()
    if not source_path.exists():
        return f"Error: File '{filename}' not found."

    normalized_format = target_format.strip().lower().lstrip(".")
    if not re.fullmatch(r"[a-z0-9]+", normalized_format):
        return f"Error: Invalid target format '{target_format}'."

    source_format = source_path.suffix.lower().lstrip(".")
    if normalized_format == source_format:
        return f"Error: Source file is already in '{source_format}' format."

    if output_dir:
        target_dir = Path(output_dir).expanduser()
    else:
        target_dir = source_path.parent
    target_path = target_dir / f"{source_path.stem}.{normalized_format}"

    converted_path = _convert_with_libreoffice(source_path, target_path)
    if converted_path is None:
        return (
            "Error converting document: LibreOffice conversion is unavailable. "
            "Install LibreOffice and ensure soffice/libreoffice is in PATH."
        )
    return _serialize_generated_file(
        converted_path, normalized_format, source_format=source_format
    )


@mcp.tool()
def list_supported_formats() -> str:
    """
    Lists all document formats supported for reading and writing.

    :return: JSON payload with readable and writable file extensions
    """
    payload = {
        "read": DocumentReaderFactory.supported_extensions(),
        "write": WRITE_FORMATS,
    }
    return json.dumps(payload, ensure_ascii=False, indent=2)


def main():
    mcp.run()


if __name__ == "__main__":
    main()
