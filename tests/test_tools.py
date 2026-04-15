"""MCP 工具函数测试。

测试内容：
- read_document MCP 工具函数测试
"""

import base64
import json
import os
from pathlib import Path
from unittest import mock

from docx import Document as DocxDocument

from mcp_documents_reader import extract_document_images, read_document

FIXTURES_DIR = Path(__file__).parent / "fixtures"
PNG_BYTES = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO7Zx6kAAAAASUVORK5CYII="
)


def create_docx_with_images(target_dir: Path, image_count: int = 2) -> Path:
    image_path = target_dir / 'sample-image.png'
    image_path.write_bytes(PNG_BYTES)

    document = DocxDocument()
    document.add_paragraph('图片测试文档')
    for _ in range(image_count):
        document.add_picture(str(image_path))

    docx_path = target_dir / 'sample-with-images.docx'
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
        self, mock_mdls: mock.MagicMock, mock_textutil: mock.MagicMock, temp_document_dir: str
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
        self, mock_mdls: mock.MagicMock, mock_command: mock.MagicMock, temp_document_dir: str
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

        assert payload['image_count'] == 2
        assert len(payload['images']) == 2
        assert payload['output_dir']
        assert payload['output_dir_uri']

        first_image = payload['images'][0]
        assert first_image['content_type'] == 'image/png'
        assert first_image['width_px'] == 1
        assert first_image['height_px'] == 1
        assert first_image['saved_path']
        assert first_image['saved_uri']
        assert Path(first_image['saved_path']).exists()

    def test_extract_document_images_with_custom_output_dir(
        self, temp_document_dir: str
    ) -> None:
        target_dir = Path(temp_document_dir)
        docx_path = create_docx_with_images(target_dir)
        export_dir = target_dir / 'exports'

        result = extract_document_images(str(docx_path), str(export_dir))
        payload = json.loads(result)

        assert payload['output_dir'] == str(export_dir.resolve())
        for image in payload['images']:
            saved_path = Path(image['saved_path'])
            assert saved_path.exists()
            assert saved_path.parent == export_dir.resolve()

    def test_extract_document_images_empty_docx(self) -> None:
        file_path = FIXTURES_DIR / 'empty.docx'

        result = extract_document_images(str(file_path))
        payload = json.loads(result)

        assert payload['image_count'] == 0
        assert payload['images'] == []
        assert payload['output_dir'] is None
        assert payload['output_dir_uri'] is None

    def test_extract_document_images_file_not_found(self) -> None:
        result = extract_document_images('nonexistent.docx')

        assert 'Error:' in result
        assert 'not found' in result

    def test_extract_document_images_unsupported_type(self) -> None:
        result = extract_document_images(str(FIXTURES_DIR / 'sample.txt'))

        assert 'Error:' in result
        assert 'DOCX files only' in result
