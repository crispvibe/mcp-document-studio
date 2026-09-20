"""MCP 工具函数测试。

测试内容：
- read_document MCP 工具函数测试
"""

import base64
import json
import os
import zipfile
from pathlib import Path
from unittest import mock

from docx import Document as DocxDocument

from mcp_documents_reader import (
    convert_document,
    extract_document_images,
    list_supported_formats,
    read_document,
)

FIXTURES_DIR = Path(__file__).parent / "fixtures"
PNG_BYTES = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO7Zx6kAAAAASUVORK5CYII="
)


def create_docx_with_images(target_dir: Path, image_count: int = 2) -> Path:
    image_path = target_dir / "sample-image.png"
    image_path.write_bytes(PNG_BYTES)

    document = DocxDocument()
    document.add_paragraph("图片测试文档")
    for _ in range(image_count):
        document.add_picture(str(image_path))

    docx_path = target_dir / "sample-with-images.docx"
    document.save(str(docx_path))
    return docx_path


class TestReadDocument:
    """read_document MCP 工具函数测试类。"""

    def test_read_document_txt_file(self) -> None:
        """测试读取 TXT 文档。"""
        file_path = FIXTURES_DIR / "sample.txt"
        result = read_document(str(file_path))

        assert "测试文本文件" in result
        assert "多行内容" in result

    def test_read_document_docx_file(self) -> None:
        """测试读取 DOCX 文档。"""
        file_path = FIXTURES_DIR / "sample.docx"
        result = read_document(str(file_path))

        assert "测试文档标题" in result

    def test_read_document_docx_file_with_images(self, temp_document_dir: str) -> None:
        docx_path = create_docx_with_images(Path(temp_document_dir))

        result = read_document(str(docx_path))

        assert "图片测试文档" in result
        assert "=== Extracted Images ===" in result
        assert "Temporary image directory:" in result
        assert "001_image1.png" in result
        assert "path:" in result
        assert "uri:" in result

    def test_read_document_pdf_file(self) -> None:
        """测试读取 PDF 文档。"""
        file_path = FIXTURES_DIR / "sample.pdf"
        result = read_document(str(file_path))

        assert "test PDF document" in result

    def test_read_document_excel_file(self) -> None:
        """测试读取 Excel 文档。"""
        file_path = FIXTURES_DIR / "sample.xlsx"
        result = read_document(str(file_path))

        assert "Sheet" in result
        assert "姓名" in result

    def test_read_document_csv_file(self, sample_csv_file: Path) -> None:
        result = read_document(str(sample_csv_file))

        assert "姓名\t年龄" in result
        assert "张三\t18" in result

    def test_read_document_markdown_file(self, sample_markdown_file: Path) -> None:
        result = read_document(str(sample_markdown_file))

        assert "# 标题" in result
        assert "- 第一项" in result

    def test_read_document_pptx_file(self, sample_pptx_file: Path) -> None:
        result = read_document(str(sample_pptx_file))

        assert "=== Slide: 1 ===" in result
        assert "演示文稿标题" in result

    def test_read_document_epub_file(self, sample_epub_file: Path) -> None:
        result = read_document(str(sample_epub_file))

        assert "=== Section: 1 ===" in result
        assert "这是 EPUB 正文。" in result

    def test_read_document_file_not_found(self) -> None:
        """测试读取不存在的文件。"""
        result = read_document("nonexistent.txt")

        assert "Error:" in result
        assert "not found" in result

    def test_read_document_unsupported_type(self) -> None:
        """测试读取不支持的文件类型。"""
        unsupported_file = FIXTURES_DIR / "test.unsupported"
        unsupported_file.write_text("test content")

        try:
            result = read_document(str(unsupported_file))

            assert "Error:" in result
            assert "Unsupported document type" in result
        finally:
            unsupported_file.unlink()

    def test_read_document_empty_file(self) -> None:
        """测试读取空文件。"""
        file_path = FIXTURES_DIR / "empty.txt"
        result = read_document(str(file_path))

        assert "No text found" in result

    def test_read_document_with_corrupted_file(self) -> None:
        """测试读取损坏的文件。"""
        file_path = FIXTURES_DIR / "corrupted.docx"
        result = read_document(str(file_path))

        assert "Error reading DOCX" in result

    def test_read_document_with_gbk_encoding(self) -> None:
        """测试读取 GBK 编码的文件。"""
        file_path = FIXTURES_DIR / "sample_gbk.txt"
        result = read_document(str(file_path))

        assert "GBK 编码" in result
        assert "中文内容" in result

    def test_read_document_with_special_characters_in_filename(
        self, temp_document_dir: str
    ) -> None:
        """测试文件名包含特殊字符。"""
        special_file = Path(temp_document_dir) / "test file (1).txt"
        special_file.write_text("special content", encoding="utf-8")

        result = read_document(str(special_file))

        assert "special content" in result

    @mock.patch("mcp_documents_reader._extract_text_with_textutil")
    @mock.patch("mcp_documents_reader._extract_text_with_mdls")
    def test_read_document_doc_file_with_mocked_extractor(
        self,
        mock_mdls: mock.MagicMock,
        mock_textutil: mock.MagicMock,
        temp_document_dir: str,
    ) -> None:
        file_path = Path(temp_document_dir) / "legacy.doc"
        file_path.write_bytes(b"fake-doc")
        mock_mdls.return_value = None
        mock_textutil.return_value = "doc content"

        result = read_document(str(file_path))

        assert result == "doc content"

    @mock.patch("mcp_documents_reader._extract_text_with_command")
    @mock.patch("mcp_documents_reader._extract_text_with_mdls")
    def test_read_document_ppt_file_with_mocked_extractor(
        self,
        mock_mdls: mock.MagicMock,
        mock_command: mock.MagicMock,
        temp_document_dir: str,
    ) -> None:
        file_path = Path(temp_document_dir) / "slides.ppt"
        file_path.write_bytes(b"fake-ppt")
        mock_mdls.return_value = None
        mock_command.return_value = "ppt content"

        result = read_document(str(file_path))

        assert result == "ppt content"


class TestReadDocumentWithPathTypes:
    """测试不同路径类型的 read_document。"""

    def test_read_document_with_absolute_path(self) -> None:
        """测试使用绝对路径读取文件。"""
        file_path = FIXTURES_DIR / "sample.txt"
        absolute_path = file_path.resolve()

        result = read_document(str(absolute_path))

        assert "测试文本文件" in result

    def test_read_document_with_relative_path(self) -> None:
        """测试使用相对路径读取文件。"""
        original_cwd = os.getcwd()
        try:
            os.chdir(FIXTURES_DIR)
            result = read_document("sample.txt")

            assert "测试文本文件" in result
        finally:
            os.chdir(original_cwd)

    def test_read_document_with_path_object(self) -> None:
        """测试使用 Path 对象读取文件。"""
        file_path = FIXTURES_DIR / "sample.txt"
        result = read_document(str(file_path))

        assert "测试文本文件" in result


class TestReadDocumentWithMockedFilesystem:
    """使用 mock 文件系统的 read_document 测试类。"""

    @mock.patch("mcp_documents_reader.Path.exists")
    @mock.patch("mcp_documents_reader.DocumentReaderFactory.is_supported")
    @mock.patch("mcp_documents_reader.DocumentReaderFactory.get_reader")
    def test_read_document_calls_reader_correctly(
        self,
        mock_get_reader: mock.MagicMock,
        mock_is_supported: mock.MagicMock,
        mock_exists: mock.MagicMock,
    ) -> None:
        """测试 read_document 正确调用 Reader。"""
        mock_exists.return_value = True
        mock_is_supported.return_value = True
        mock_reader = mock.MagicMock()
        mock_reader.read.return_value = "test content"
        mock_get_reader.return_value = mock_reader

        result = read_document("test.txt")

        mock_exists.assert_called_once()
        mock_is_supported.assert_called_once()
        mock_get_reader.assert_called_once()
        mock_reader.read.assert_called_once()

        assert result == "test content"

    @mock.patch("mcp_documents_reader.Path.exists")
    def test_read_document_file_not_exists_mock(
        self, mock_exists: mock.MagicMock
    ) -> None:
        """测试文件不存在的情况（使用 mock）。"""
        mock_exists.return_value = False

        result = read_document("test.txt")

        assert "Error:" in result
        assert "not found" in result

    @mock.patch("mcp_documents_reader.Path.exists")
    @mock.patch("mcp_documents_reader.DocumentReaderFactory.is_supported")
    def test_read_document_unsupported_type_mock(
        self, mock_is_supported: mock.MagicMock, mock_exists: mock.MagicMock
    ) -> None:
        """测试不支持的文件类型（使用 mock）。"""
        mock_exists.return_value = True
        mock_is_supported.return_value = False

        result = read_document("test.xyz")

        assert "Error:" in result
        assert "Unsupported document type" in result

    @mock.patch("mcp_documents_reader.Path.exists")
    @mock.patch("mcp_documents_reader.DocumentReaderFactory.is_supported")
    @mock.patch("mcp_documents_reader.DocumentReaderFactory.get_reader")
    def test_read_document_reader_exception(
        self,
        mock_get_reader: mock.MagicMock,
        mock_is_supported: mock.MagicMock,
        mock_exists: mock.MagicMock,
    ) -> None:
        """测试 Reader 抛出异常的情况。"""
        mock_exists.return_value = True
        mock_is_supported.return_value = True
        mock_get_reader.side_effect = Exception("Reader error")

        result = read_document("test.txt")

        assert "Error reading document" in result


class TestExtractDocumentImages:
    def test_extract_document_images_docx_file(self, temp_document_dir: str) -> None:
        docx_path = create_docx_with_images(Path(temp_document_dir))

        result = extract_document_images(str(docx_path))
        payload = json.loads(result)

        assert payload["image_count"] == 2
        assert len(payload["images"]) == 2
        assert payload["output_dir"]
        assert payload["output_dir_uri"]

        first_image = payload["images"][0]
        assert first_image["content_type"] == "image/png"
        assert first_image["width_px"] == 1
        assert first_image["height_px"] == 1
        assert first_image["saved_path"]
        assert first_image["saved_uri"]
        assert Path(first_image["saved_path"]).exists()

    def test_extract_document_images_with_custom_output_dir(
        self, temp_document_dir: str
    ) -> None:
        target_dir = Path(temp_document_dir)
        docx_path = create_docx_with_images(target_dir)
        export_dir = target_dir / "exports"

        result = extract_document_images(str(docx_path), str(export_dir))
        payload = json.loads(result)

        assert payload["output_dir"] == str(export_dir.resolve())
        for image in payload["images"]:
            saved_path = Path(image["saved_path"])
            assert saved_path.exists()
            assert saved_path.parent == export_dir.resolve()

    def test_extract_document_images_empty_docx(self) -> None:
        file_path = FIXTURES_DIR / "empty.docx"

        result = extract_document_images(str(file_path))
        payload = json.loads(result)

        assert payload["image_count"] == 0
        assert payload["images"] == []
        assert payload["output_dir"] is None
        assert payload["output_dir_uri"] is None

    def test_extract_document_images_file_not_found(self) -> None:
        result = extract_document_images("nonexistent.docx")

        assert "Error:" in result
        assert "not found" in result

    def test_extract_document_images_unsupported_type(self) -> None:
        result = extract_document_images(str(FIXTURES_DIR / "sample.txt"))

        assert "Error:" in result
        assert "DOCX files only" in result


class TestReadDocumentNewFormats:
    """新增格式（HTML/JSON/XML/YAML）的 read_document 测试。"""

    def test_read_document_html_file(self, temp_document_dir: str) -> None:
        html_path = Path(temp_document_dir) / "page.html"
        html_path.write_text(
            "<html><body><h1>网页标题</h1><p>网页正文</p></body></html>",
            encoding="utf-8",
        )

        result = read_document(str(html_path))

        assert "网页标题" in result
        assert "网页正文" in result

    def test_read_document_json_file(self, temp_document_dir: str) -> None:
        json_path = Path(temp_document_dir) / "data.json"
        json_path.write_text('{"key":"值"}', encoding="utf-8")

        result = read_document(str(json_path))

        assert '"key": "值"' in result

    def test_read_document_xml_file(self, temp_document_dir: str) -> None:
        xml_path = Path(temp_document_dir) / "config.xml"
        xml_path.write_text("<config><name>测试</name></config>", encoding="utf-8")

        result = read_document(str(xml_path))

        assert "<config>" in result
        assert "测试" in result

    def test_read_document_yaml_file(self, temp_document_dir: str) -> None:
        yaml_path = Path(temp_document_dir) / "config.yaml"
        yaml_path.write_text("key: 配置值\n", encoding="utf-8")

        result = read_document(str(yaml_path))

        assert "key: 配置值" in result

    def test_read_document_expanduser(self, temp_document_dir: str) -> None:
        txt_path = Path(temp_document_dir) / "home.txt"
        txt_path.write_text("home content", encoding="utf-8")

        with mock.patch.object(
            Path, "expanduser", autospec=True, return_value=txt_path
        ):
            result = read_document("~/home.txt")

        assert "home content" in result


class TestConvertDocument:
    """convert_document MCP 工具函数测试类。"""

    @mock.patch("mcp_documents_reader._convert_with_libreoffice")
    def test_convert_document_success(
        self, mock_convert: mock.MagicMock, temp_document_dir: str
    ) -> None:
        source_path = Path(temp_document_dir) / "report.docx"
        source_path.write_bytes(b"fake-docx")
        target_path = Path(temp_document_dir) / "report.pdf"
        target_path.write_bytes(b"fake-pdf")
        mock_convert.return_value = target_path

        result = convert_document(str(source_path), "pdf")

        payload = json.loads(result)
        assert payload["format"] == "pdf"
        assert payload["source_format"] == "docx"
        mock_convert.assert_called_once()
        call_args = mock_convert.call_args[0]
        assert call_args[1] == target_path

    @mock.patch("mcp_documents_reader._convert_with_libreoffice")
    def test_convert_document_custom_output_dir(
        self, mock_convert: mock.MagicMock, temp_document_dir: str
    ) -> None:
        source_path = Path(temp_document_dir) / "report.docx"
        source_path.write_bytes(b"fake-docx")
        output_dir = Path(temp_document_dir) / "out"
        mock_convert.return_value = output_dir / "report.pdf"

        result = convert_document(str(source_path), "pdf", str(output_dir))

        payload = json.loads(result)
        assert payload["format"] == "pdf"
        assert mock_convert.call_args[0][1].parent == output_dir

    def test_convert_document_file_not_found(self, temp_document_dir: str) -> None:
        missing = Path(temp_document_dir) / "missing.docx"
        result = convert_document(str(missing), "pdf")

        assert "not found" in result

    def test_convert_document_invalid_format(self, temp_document_dir: str) -> None:
        source_path = Path(temp_document_dir) / "report.docx"
        source_path.write_bytes(b"fake-docx")

        result = convert_document(str(source_path), "../evil")

        assert "Invalid target format" in result

    def test_convert_document_same_format(self, temp_document_dir: str) -> None:
        source_path = Path(temp_document_dir) / "report.docx"
        source_path.write_bytes(b"fake-docx")

        result = convert_document(str(source_path), ".docx")

        assert "already" in result

    @mock.patch("mcp_documents_reader._convert_with_libreoffice", return_value=None)
    def test_convert_document_libreoffice_unavailable(
        self, mock_convert: mock.MagicMock, temp_document_dir: str
    ) -> None:
        source_path = Path(temp_document_dir) / "report.docx"
        source_path.write_bytes(b"fake-docx")

        result = convert_document(str(source_path), "pdf")

        assert "LibreOffice conversion is unavailable" in result
        mock_convert.assert_called_once()


class TestListSupportedFormats:
    """list_supported_formats MCP 工具函数测试类。"""

    def test_list_supported_formats(self) -> None:
        result = list_supported_formats()
        payload = json.loads(result)

        assert ".docx" in payload["read"]
        assert ".pdf" in payload["read"]
        assert ".html" in payload["read"]
        assert ".json" in payload["read"]
        assert ".docx" in payload["write"]["word"]
        assert ".pptx" in payload["write"]["presentation"]
        assert ".xlsx" in payload["write"]["spreadsheet"]


def _fake_png(size: int = 300) -> bytes:
    """伪造足够大的 PNG blob（签名 + 填充 + IEND），用于签名扫描。"""
    return b"\x89PNG\r\n\x1a\n" + b"\x00" * (size - 20) + b"IEND" + b"\x00" * 8


class TestReadDocumentImages:
    """read_document_images 工具测试。"""

    def test_returns_image_blocks_for_docx(self, tmp_path: Path) -> None:
        from mcp_documents_reader import read_document_images

        docx_path = tmp_path / "with_image.docx"
        with zipfile.ZipFile(docx_path, "w") as archive:
            archive.writestr("[Content_Types].xml", "<xml/>")
            archive.writestr("word/media/image1.png", _fake_png())
            archive.writestr(
                "word/media/photo.jpg", b"\xff\xd8\xff" + b"\x00" * 300 + b"\xff\xd9"
            )

        result = read_document_images(str(docx_path))

        assert isinstance(result, list)
        assert "Found 2 embedded image(s)" in result[0]
        image_blocks = [b for b in result if getattr(b, "data", None)]
        assert len(image_blocks) == 2

    def test_legacy_doc_signature_scan(self, tmp_path: Path) -> None:
        from mcp_documents_reader import read_document_images

        doc_path = tmp_path / "old.doc"
        doc_path.write_bytes(
            b"\xd0\xcf\x11\xe0" + b"\x00" * 100 + _fake_png() + b"\x00" * 50
        )

        result = read_document_images(str(doc_path))

        assert "Found 1 embedded image(s)" in result[0]
        image_blocks = [b for b in result if getattr(b, "data", None)]
        assert len(image_blocks) == 1

    def test_epub_images(self, tmp_path: Path) -> None:
        from mcp_documents_reader import read_document_images

        epub_path = tmp_path / "book.epub"
        with zipfile.ZipFile(epub_path, "w") as archive:
            archive.writestr("OEBPS/cover.png", _fake_png())
            archive.writestr("OEBPS/text.xhtml", "<p>hi</p>")

        result = read_document_images(str(epub_path))

        assert "Found 1 embedded image(s)" in result[0]

    def test_max_images_cap(self, tmp_path: Path) -> None:
        from mcp_documents_reader import read_document_images

        xlsx_path = tmp_path / "book.xlsx"
        with zipfile.ZipFile(xlsx_path, "w") as archive:
            for index in range(5):
                archive.writestr(f"xl/media/img{index}.png", _fake_png())

        result = read_document_images(str(xlsx_path), max_images=2)

        assert "showing first 2" in result[0]
        image_blocks = [b for b in result if getattr(b, "data", None)]
        assert len(image_blocks) == 2

    def test_unsupported_extension(self, tmp_path: Path) -> None:
        from mcp_documents_reader import read_document_images

        txt_path = tmp_path / "file.txt"
        txt_path.write_text("hello")

        result = read_document_images(str(txt_path))

        assert result[0].startswith("Error:")

    def test_missing_file(self, tmp_path: Path) -> None:
        from mcp_documents_reader import read_document_images

        result = read_document_images(str(tmp_path / "nope.docx"))

        assert "not found" in result[0]

    def test_no_images_found(self, tmp_path: Path) -> None:
        from mcp_documents_reader import read_document_images

        docx_path = tmp_path / "plain.docx"
        with zipfile.ZipFile(docx_path, "w") as archive:
            archive.writestr("[Content_Types].xml", "<xml/>")

        result = read_document_images(str(docx_path))

        assert "No embedded images" in result[0]

    def test_pdf_page_images(self, tmp_path: Path) -> None:
        from mcp_documents_reader import read_document_images

        pdf_path = tmp_path / "scan.pdf"
        pdf_path.write_bytes(b"%PDF-1.4 fake")

        fake_image = mock.MagicMock()
        fake_image.name = "page_scan.png"
        fake_image.data = _fake_png()
        fake_page = mock.MagicMock()
        fake_page.images = [fake_image]
        fake_reader = mock.MagicMock()
        fake_reader.is_encrypted = False
        fake_reader.pages = [fake_page]

        with mock.patch("mcp_documents_reader.PyPdfReader", return_value=fake_reader):
            result = read_document_images(str(pdf_path))

        assert "Found 1 embedded image(s)" in result[0]
        assert "page1_page_scan.png" in result[1]

    def test_corrupt_container_returns_no_images(self, tmp_path: Path) -> None:
        from mcp_documents_reader import read_document_images

        bad_path = tmp_path / "corrupt.docx"
        bad_path.write_bytes(b"not a zip at all")

        result = read_document_images(str(bad_path))

        assert "No embedded images" in result[0]


class TestIterEmbeddedImageBlobs:
    def test_skips_tiny_blobs(self) -> None:
        from mcp_documents_reader import _iter_embedded_image_blobs

        tiny = b"\x89PNG\r\n\x1a\n" + b"IEND" + b"\x00" * 8
        assert list(_iter_embedded_image_blobs(tiny)) == []

    def test_unterminated_blob_runs_to_eof(self) -> None:
        from mcp_documents_reader import _iter_embedded_image_blobs

        data = b"\xff\xd8\xff" + b"\x00" * 400  # no EOI marker
        blobs = list(_iter_embedded_image_blobs(data))
        assert len(blobs) == 1
        assert blobs[0][0] == "jpeg"
