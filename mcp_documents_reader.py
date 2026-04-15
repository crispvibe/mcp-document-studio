import json
import os
import csv
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from abc import ABC, abstractmethod
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree

from docx import Document as DocxDocument
from mcp.server.fastmcp import FastMCP
from openpyxl import load_workbook
from pypdf import PdfReader as PyPdfReader
from typing_extensions import override

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
            f"[{index}] {filename} | {width_px}x{height_px} | path: {saved_path} | uri: {saved_uri}"
        )

    if text_content:
        return text_content + "\n" + "\n".join(image_lines)
    return "\n".join(image_lines).strip()


def _read_text_file(
    file_path: str, *, empty_message: str, error_prefix: str
) -> str:
    for encoding in TEXT_ENCODINGS:
        try:
            with open(file_path, "r", encoding=encoding, newline="") as file:
                text = file.read()
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
        manifest[item_id] = str((opf_parent / href).as_posix())

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
) -> DocxDocument:
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
            "python-pptx is required for presentation generation. Install project dependencies first."
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
        if title_slide.shapes.title is not None:
            title_slide.shapes.title.text = title or ""
        if len(title_slide.placeholders) > 1:
            title_slide.placeholders[1].text = subtitle or ""

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


class DocxReader(DocumentReader):
    """DOCX document reader implementation"""

    @override
    def read(self, file_path: str) -> str:
        """Read and extract text from DOCX file"""
        try:
            doc = DocxDocument(file_path)
            text = []

            for paragraph in doc.paragraphs:
                if paragraph.text:
                    text.append(paragraph.text)

            for table in doc.tables:
                for row in table.rows:
                    row_text = []
                    for cell in row.cells:
                        cell_text = " ".join([p.text for p in cell.paragraphs]).strip()
                        if cell_text:
                            row_text.append(cell_text)
                    if row_text:
                        text.append("\t".join(row_text))

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
            if "image" not in rel.reltype:
                continue

            image_part = rel.target_part
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
            blob = image_info.pop("blob")
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
                text = []

                for page in pdf_reader.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text.append(page_text.strip())

                extracted_text = "\n\n".join(text)
                return extracted_text if extracted_text else "No text found in the PDF."
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
                rows: list[str] = []
                with open(file_path, "r", encoding=encoding, newline="") as file:
                    reader = csv.reader(file)
                    for row in reader:
                        if not row:
                            continue
                        row_text = [cell.strip() for cell in row]
                        if any(row_text):
                            rows.append("\t".join(row_text))
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
            )
            if extracted:
                return extracted
            return (
                "Error reading DOC: No available extractor succeeded. "
                "Try installing antiword or LibreOffice on this machine."
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
            )
            if extracted:
                return extracted
            return (
                "Error reading PPT: No available extractor succeeded. "
                "Try installing catdoc/catppt or LibreOffice on this machine."
            )
        except Exception as exc:
            return f"Error reading PPT: {str(exc)}"


class PptxReader(DocumentReader):
    @override
    def read(self, file_path: str) -> str:
        try:
            with zipfile.ZipFile(file_path) as archive:
                slide_names = [
                    name
                    for name in archive.namelist()
                    if PPTX_SLIDE_PATTERN.match(name) is not None
                ]
                ordered_slide_names = sorted(
                    slide_names,
                    key=lambda name: int(PPTX_SLIDE_PATTERN.match(name).group(1)),
                )
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


class ExcelReader(DocumentReader):
    """Excel document reader implementation"""

    @override
    def read(self, file_path: str) -> str:
        """Read and extract text from Excel file"""
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
    file_path = Path(filename)

    if not file_path.exists():
        return f"Error: File '{filename}' not found."

    if not DocumentReaderFactory.is_supported(str(file_path)):
        return f"Error: Unsupported document type for file '{filename}'."

    try:
        reader = DocumentReaderFactory.get_reader(str(file_path))
        content = reader.read(str(file_path))
        if isinstance(reader, DocxReader) and not content.startswith("Error reading DOCX:"):
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
    file_path = Path(filename)

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

        with tempfile.TemporaryDirectory(prefix="mcp-document-writer-word-") as temp_dir:
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
    Generates a PowerPoint presentation in PPTX format, or PPT via LibreOffice conversion.

    :param filename: Target output path ending with .pptx or .ppt
    :param title: Optional title slide title
    :param subtitle: Optional title slide subtitle
    :param slides: Optional slide definitions containing title, paragraphs, bullets, and table
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


def main():
    mcp.run()


if __name__ == "__main__":
    main()
