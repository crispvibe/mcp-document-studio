"""DocumentReader 抽象基类及各 Reader 实现测试。

测试内容：
- DocumentReader 抽象基类验证
- DocxReader 单元测试
- PdfReader 单元测试
- TxtReader 单元测试（多种编码）
- ExcelReader 单元测试
"""

import zipfile
from pathlib import Path
from textwrap import dedent
from unittest import mock

import pytest

from mcp_documents_reader import (
    CsvReader,
    DocReader,
    DocumentReader,
    DocxReader,
    EpubReader,
    ExcelReader,
    HtmlReader,
    JsonReader,
    MarkdownReader,
    OdfReader,
    PdfReader,
    PptReader,
    PptxReader,
    RtfReader,
    TxtReader,
)


class TestDocumentReader:
    """DocumentReader 抽象基类测试类。"""

    def test_cannot_instantiate_abstract_class(self) -> None:
        """测试无法直接实例化抽象基类。"""
        with pytest.raises(TypeError):
            DocumentReader()  # type: ignore[abstract]

    def test_subclass_must_implement_read(self) -> None:
        """测试子类必须实现 read 方法。"""

        class IncompleteReader(DocumentReader):
            """不完整的 Reader 实现。"""

            pass

        with pytest.raises(TypeError):
            IncompleteReader()  # type: ignore[abstract]

    def test_subclass_with_read_implementation(self) -> None:
        """测试正确实现 read 方法的子类可以实例化。"""

        class CompleteReader(DocumentReader):
            """完整的 Reader 实现。"""

            def read(self, file_path: str) -> str:  # noqa: ARG002
                """读取文件内容。"""
                return "test content"

        reader = CompleteReader()
        assert reader.read("test.txt") == "test content"


class TestDocxReader:
    """DocxReader 测试类。"""

    def test_read_docx_with_content(self, sample_docx_file: Path) -> None:
        """测试读取包含内容的 DOCX 文件。

        Args:
            sample_docx_file: 示例 DOCX 文件路径
        """
        reader = DocxReader()
        result = reader.read(str(sample_docx_file))

        # 验证提取的内容包含预期文本
        assert "测试文档标题" in result
        assert "第一段文本内容" in result
        assert "第二段文本内容" in result

    def test_read_docx_with_table(self, sample_docx_file: Path) -> None:
        """测试读取包含表格的 DOCX 文件。

        Args:
            sample_docx_file: 示例 DOCX 文件路径
        """
        reader = DocxReader()
        result = reader.read(str(sample_docx_file))

        # 验证表格内容被提取
        assert "姓名" in result
        assert "年龄" in result
        assert "张三" in result
        assert "李四" in result

    def test_read_empty_docx(self, empty_docx_file: Path) -> None:
        """测试读取空的 DOCX 文件。

        Args:
            empty_docx_file: 空 DOCX 文件路径
        """
        reader = DocxReader()
        result = reader.read(str(empty_docx_file))

        # 空文档应返回提示信息
        assert "No text found" in result

    def test_read_corrupted_docx(self, fixtures_dir: Path) -> None:
        """测试读取损坏的 DOCX 文件。

        Args:
            fixtures_dir: fixtures 目录路径
        """
        reader = DocxReader()
        corrupted_file = fixtures_dir / "corrupted.docx"
        result = reader.read(str(corrupted_file))

        # 损坏文件应返回错误信息
        assert "Error reading DOCX" in result

    def test_read_nonexistent_file(self) -> None:
        """测试读取不存在的文件。"""
        reader = DocxReader()
        result = reader.read("/nonexistent/path/file.docx")

        assert "Error reading DOCX" in result

    def test_read_docx_preserves_block_order(self, temp_document_dir: str) -> None:
        """测试段落和表格按文档顺序提取，而不是全部表格排在末尾。"""
        from docx import Document as DocxDocument

        docx_path = Path(temp_document_dir) / "ordered.docx"
        document = DocxDocument()
        document.add_paragraph("before-table")
        table = document.add_table(rows=1, cols=1)
        table.rows[0].cells[0].text = "in-table"
        document.add_paragraph("after-table")
        document.save(str(docx_path))

        reader = DocxReader()
        result = reader.read(str(docx_path))

        before_pos = result.index("before-table")
        table_pos = result.index("in-table")
        after_pos = result.index("after-table")
        assert before_pos < table_pos < after_pos


class TestPdfReader:
    """PdfReader 测试类。"""

    def test_read_pdf_with_content(self, sample_pdf_file: Path) -> None:
        """测试读取包含内容的 PDF 文件。

        Args:
            sample_pdf_file: 示例 PDF 文件路径
        """
        reader = PdfReader()
        result = reader.read(str(sample_pdf_file))

        # 验证提取的内容包含预期文本
        assert "test PDF document" in result
        assert "multiple lines" in result

    def test_read_empty_pdf(self, empty_pdf_file: Path) -> None:
        """测试读取空的 PDF 文件。

        Args:
            empty_pdf_file: 空 PDF 文件路径
        """
        reader = PdfReader()
        result = reader.read(str(empty_pdf_file))

        # 空 PDF 应返回提示信息
        assert "No text found" in result

    def test_read_corrupted_pdf(self, fixtures_dir: Path) -> None:
        """测试读取损坏的 PDF 文件。

        Args:
            fixtures_dir: fixtures 目录路径
        """
        reader = PdfReader()
        corrupted_file = fixtures_dir / "corrupted.pdf"
        result = reader.read(str(corrupted_file))

        # 损坏文件应返回错误信息
        assert "Error reading PDF" in result

    def test_read_nonexistent_file(self) -> None:
        """测试读取不存在的文件。"""
        reader = PdfReader()
        result = reader.read("/nonexistent/path/file.pdf")

        assert "Error reading PDF" in result

    def test_read_encrypted_pdf(self, temp_document_dir: str) -> None:
        """测试加密 PDF 会尝试空密码解密后读取。"""
        from pypdf import PdfWriter

        pdf_path = Path(temp_document_dir) / "encrypted.pdf"
        writer = PdfWriter()
        writer.add_blank_page(width=72, height=72)
        writer.encrypt("")
        with open(pdf_path, "wb") as file:
            writer.write(file)

        reader = PdfReader()
        result = reader.read(str(pdf_path))

        assert "Error reading PDF" not in result


class TestTxtReader:
    """TxtReader 测试类。"""

    def test_read_utf8_txt(self, sample_txt_file: Path) -> None:
        """测试读取 UTF-8 编码的 TXT 文件。

        Args:
            sample_txt_file: 示例 TXT 文件路径
        """
        reader = TxtReader()
        result = reader.read(str(sample_txt_file))

        # 验证提取的内容
        assert "测试文本文件" in result
        assert "多行内容" in result
        assert "中文" in result

    def test_read_gbk_txt(self, sample_txt_gbk_file: Path) -> None:
        """测试读取 GBK 编码的 TXT 文件。

        Args:
            sample_txt_gbk_file: GBK 编码的示例 TXT 文件路径
        """
        reader = TxtReader()
        result = reader.read(str(sample_txt_gbk_file))

        # 验证 GBK 编码文件被正确读取
        assert "GBK 编码" in result
        assert "中文内容" in result

    def test_read_empty_txt(self, empty_txt_file: Path) -> None:
        """测试读取空的 TXT 文件。

        Args:
            empty_txt_file: 空 TXT 文件路径
        """
        reader = TxtReader()
        result = reader.read(str(empty_txt_file))

        # 空文件应返回提示信息
        assert "No text found" in result

    def test_read_binary_file(self, fixtures_dir: Path) -> None:
        """测试读取二进制文件。

        注意：latin-1 编码可以解码任何字节序列，所以二进制文件
        会被 latin-1 成功读取，返回解码后的内容。

        Args:
            fixtures_dir: fixtures 目录路径
        """
        reader = TxtReader()
        binary_file = fixtures_dir / "binary.txt"
        result = reader.read(str(binary_file))

        # latin-1 可以解码任何字节，所以会返回解码后的内容
        # 验证返回的是解码后的内容，而不是错误信息
        assert result != ""
        assert "Error reading TXT" not in result

    def test_read_nonexistent_file(self) -> None:
        """测试读取不存在的文件。"""
        reader = TxtReader()
        result = reader.read("/nonexistent/path/file.txt")

        # 不存在的文件应返回错误信息
        assert "Error reading TXT" in result


class TestExcelReader:
    """ExcelReader 测试类。"""

    def test_read_excel_with_content(self, sample_excel_file: Path) -> None:
        """测试读取包含内容的 Excel 文件。

        Args:
            sample_excel_file: 示例 Excel 文件路径
        """
        reader = ExcelReader()
        result = reader.read(str(sample_excel_file))

        # 验证提取的内容包含工作表标题
        assert "=== Sheet: Sheet1 ===" in result
        assert "=== Sheet: Sheet2 ===" in result

    def test_read_excel_sheet1_data(self, sample_excel_file: Path) -> None:
        """测试读取 Excel Sheet1 数据。

        Args:
            sample_excel_file: 示例 Excel 文件路径
        """
        reader = ExcelReader()
        result = reader.read(str(sample_excel_file))

        # 验证 Sheet1 数据
        assert "姓名" in result
        assert "年龄" in result
        assert "张三" in result
        assert "李四" in result

    def test_read_excel_sheet2_data(self, sample_excel_file: Path) -> None:
        """测试读取 Excel Sheet2 数据。

        Args:
            sample_excel_file: 示例 Excel 文件路径
        """
        reader = ExcelReader()
        result = reader.read(str(sample_excel_file))

        # 验证 Sheet2 数据
        assert "产品" in result
        assert "价格" in result
        assert "苹果" in result
        assert "香蕉" in result

    def test_read_empty_excel(self, empty_excel_file: Path) -> None:
        """测试读取空的 Excel 文件。

        Args:
            empty_excel_file: 空 Excel 文件路径
        """
        reader = ExcelReader()
        result = reader.read(str(empty_excel_file))

        # 空 Excel 应返回提示信息或工作表标题
        # 由于有空工作表，可能返回工作表标题但没有数据
        assert "Sheet" in result or "No text found" in result

    def test_read_corrupted_excel(self, fixtures_dir: Path) -> None:
        """测试读取损坏的 Excel 文件。

        Args:
            fixtures_dir: fixtures 目录路径
        """
        reader = ExcelReader()
        corrupted_file = fixtures_dir / "corrupted.xlsx"
        result = reader.read(str(corrupted_file))

        # 损坏文件应返回错误信息
        assert "Error reading Excel" in result

    def test_read_nonexistent_file(self) -> None:
        """测试读取不存在的文件。"""
        reader = ExcelReader()
        result = reader.read("/nonexistent/path/file.xlsx")

        assert "Error reading Excel" in result

    @mock.patch("mcp_documents_reader._extract_text_with_command")
    def test_read_xls_via_xls2csv(
        self, mock_command: mock.MagicMock, temp_document_dir: str
    ) -> None:
        """测试 .xls 文件通过 xls2csv 回退读取。"""
        xls_path = Path(temp_document_dir) / "legacy.xls"
        xls_path.write_bytes(b"fake-xls")
        mock_command.return_value = '"姓名","年龄"\n"张三","18"\n'

        reader = ExcelReader()
        result = reader.read(str(xls_path))

        assert "姓名\t年龄" in result
        assert "张三\t18" in result
        mock_command.assert_called_once_with(str(xls_path), "xls2csv")

    @mock.patch("mcp_documents_reader._convert_with_libreoffice")
    @mock.patch("mcp_documents_reader._extract_text_with_command")
    def test_read_xls_via_libreoffice_conversion(
        self,
        mock_command: mock.MagicMock,
        mock_convert: mock.MagicMock,
        temp_document_dir: str,
        sample_excel_file: Path,
    ) -> None:
        """测试 .xls 文件通过 LibreOffice 转 xlsx 后读取。"""
        xls_path = Path(temp_document_dir) / "legacy.xls"
        xls_path.write_bytes(b"fake-xls")
        mock_command.return_value = None
        mock_convert.return_value = sample_excel_file

        reader = ExcelReader()
        result = reader.read(str(xls_path))

        assert "姓名" in result
        assert "=== Sheet: Sheet1 ===" in result
        mock_convert.assert_called_once()

    @mock.patch("mcp_documents_reader._convert_with_libreoffice", return_value=None)
    @mock.patch("mcp_documents_reader._extract_text_with_command", return_value=None)
    @mock.patch("mcp_documents_reader._extract_legacy_binary_text", return_value=None)
    def test_read_xls_no_extractor_available(
        self,
        mock_binary: mock.MagicMock,
        mock_command: mock.MagicMock,
        mock_convert: mock.MagicMock,
        temp_document_dir: str,
    ) -> None:
        """测试无可用提取器时 .xls 返回明确错误信息。"""
        xls_path = Path(temp_document_dir) / "legacy.xls"
        xls_path.write_bytes(b"\x00\x01" * 32)

        reader = ExcelReader()
        with mock.patch.object(reader, "_read_with_calamine", return_value=None):
            result = reader.read(str(xls_path))

        assert "Error reading Excel" in result
        assert "xls2csv" in result or "LibreOffice" in result


class TestCsvReader:
    def test_read_csv_with_content(self, sample_csv_file: Path) -> None:
        reader = CsvReader()
        result = reader.read(str(sample_csv_file))

        assert "姓名\t年龄" in result
        assert "张三\t18" in result
        assert "李四\t20" in result


class TestMarkdownReader:
    def test_read_markdown_with_content(self, sample_markdown_file: Path) -> None:
        reader = MarkdownReader()
        result = reader.read(str(sample_markdown_file))

        assert "# 标题" in result
        assert "- 第一项" in result


class TestPptxReader:
    def test_read_pptx_with_content(self, sample_pptx_file: Path) -> None:
        reader = PptxReader()
        result = reader.read(str(sample_pptx_file))

        assert "=== Slide: 1 ===" in result
        assert "演示文稿标题" in result
        assert "第二行内容" in result

    def test_read_pptx_respects_slide_order(self, temp_document_dir: str) -> None:
        """测试幻灯片顺序由 presentation.xml 的 sldIdLst 决定而非文件名。"""
        pptx_path = Path(temp_document_dir) / "unordered.pptx"

        def slide_xml(text: str) -> str:
            return dedent(
                f"""\
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
                  <p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld>
                </p:sld>
                """
            ).strip()

        with zipfile.ZipFile(pptx_path, "w") as archive:
            archive.writestr(
                "[Content_Types].xml",
                '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>',
            )
            archive.writestr(
                "ppt/presentation.xml",
                dedent(
                    """\
                    <?xml version="1.0" encoding="UTF-8"?>
                    <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                      <p:sldIdLst>
                        <p:sldId id="256" r:id="rId2"/>
                        <p:sldId id="257" r:id="rId1"/>
                      </p:sldIdLst>
                    </p:presentation>
                    """
                ).strip(),
            )
            archive.writestr(
                "ppt/_rels/presentation.xml.rels",
                dedent(
                    """\
                    <?xml version="1.0" encoding="UTF-8"?>
                    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
                    </Relationships>
                    """
                ).strip(),
            )
            archive.writestr("ppt/slides/slide1.xml", slide_xml("第二章内容"))
            archive.writestr("ppt/slides/slide2.xml", slide_xml("第一章内容"))

        reader = PptxReader()
        result = reader.read(str(pptx_path))

        assert result.index("第一章内容") < result.index("第二章内容")
        assert "=== Slide: 1 ===" in result
        assert "=== Slide: 2 ===" in result

    def test_read_pptx_filename_order_fallback(self, temp_document_dir: str) -> None:
        """测试缺失 presentation.xml 时回退到文件名数字排序。"""
        pptx_path = Path(temp_document_dir) / "fallback.pptx"

        def slide_xml(text: str) -> str:
            return (
                '<?xml version="1.0"?>'
                '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
                "<p:cSld><p:spTree><p:sp><p:txBody>"
                f"<a:p><a:r><a:t>{text}</a:t></a:r></a:p>"
                "</p:txBody></p:sp></p:spTree></p:cSld></p:sld>"
            )

        with zipfile.ZipFile(pptx_path, "w") as archive:
            archive.writestr("ppt/slides/slide2.xml", slide_xml("第二页"))
            archive.writestr("ppt/slides/slide1.xml", slide_xml("第一页"))

        reader = PptxReader()
        result = reader.read(str(pptx_path))

        assert result.index("第一页") < result.index("第二页")


class TestEpubReader:
    def test_read_epub_with_content(self, sample_epub_file: Path) -> None:
        reader = EpubReader()
        result = reader.read(str(sample_epub_file))

        assert "=== Section: 1 ===" in result
        assert "第一章标题" in result
        assert "这是 EPUB 正文。" in result

    def test_read_epub_with_encoded_href(self, temp_document_dir: str) -> None:
        """测试 manifest href 带 URL 编码和片段时能正确读取章节。"""
        epub_path = Path(temp_document_dir) / "encoded.epub"
        with zipfile.ZipFile(epub_path, "w") as archive:
            archive.writestr("mimetype", "application/epub+zip")
            archive.writestr(
                "META-INF/container.xml",
                dedent(
                    """\
                    <?xml version="1.0" encoding="UTF-8"?>
                    <container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
                      <rootfiles>
                        <rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/>
                      </rootfiles>
                    </container>
                    """
                ).strip(),
            )
            archive.writestr(
                "OEBPS/content.opf",
                dedent(
                    """\
                    <?xml version="1.0" encoding="UTF-8"?>
                    <package version="2.0" xmlns="http://www.idpf.org/2007/opf">
                      <manifest>
                        <item id="ch1" href="chapter%201.xhtml#start" media-type="application/xhtml+xml"/>
                      </manifest>
                      <spine>
                        <itemref idref="ch1"/>
                      </spine>
                    </package>
                    """
                ).strip(),
            )
            archive.writestr(
                "OEBPS/chapter 1.xhtml",
                '<html xmlns="http://www.w3.org/1999/xhtml"><body><p>编码章节内容</p></body></html>',
            )

        reader = EpubReader()
        result = reader.read(str(epub_path))

        assert "编码章节内容" in result


class TestDocReader:
    @mock.patch("mcp_documents_reader._extract_text_with_libreoffice")
    @mock.patch("mcp_documents_reader._extract_text_with_command")
    @mock.patch("mcp_documents_reader._extract_text_with_textutil")
    @mock.patch("mcp_documents_reader._extract_text_with_mdls")
    def test_read_doc_uses_available_fallback(
        self,
        mock_mdls: mock.MagicMock,
        mock_textutil: mock.MagicMock,
        mock_command: mock.MagicMock,
        mock_libreoffice: mock.MagicMock,
    ) -> None:
        mock_mdls.return_value = None
        mock_textutil.return_value = "legacy doc content"
        mock_command.return_value = None
        mock_libreoffice.return_value = None

        reader = DocReader()
        result = reader.read("legacy.doc")

        assert result == "legacy doc content"
        mock_textutil.assert_called_once_with("legacy.doc")


class TestPptReader:
    @mock.patch("mcp_documents_reader._extract_text_with_libreoffice")
    @mock.patch("mcp_documents_reader._extract_text_with_command")
    @mock.patch("mcp_documents_reader._extract_text_with_mdls")
    def test_read_ppt_uses_available_fallback(
        self,
        mock_mdls: mock.MagicMock,
        mock_command: mock.MagicMock,
        mock_libreoffice: mock.MagicMock,
    ) -> None:
        mock_mdls.return_value = None
        mock_command.return_value = "legacy ppt content"
        mock_libreoffice.return_value = None

        reader = PptReader()
        result = reader.read("slides.ppt")

        assert result == "legacy ppt content"
        mock_command.assert_called_once_with("slides.ppt", "catppt")


class TestHtmlReader:
    def test_read_html_with_content(self, temp_document_dir: str) -> None:
        html_path = Path(temp_document_dir) / "page.html"
        html_path.write_text(
            dedent(
                """\
                <html><head><title>标题</title><style>body{color:red}</style></head>
                <body><h1>主标题</h1><p>正文段落</p><script>var x=1;</script>
                <ul><li>列表项</li></ul></body></html>
                """
            ),
            encoding="utf-8",
        )

        reader = HtmlReader()
        result = reader.read(str(html_path))

        assert "主标题" in result
        assert "正文段落" in result
        assert "列表项" in result
        assert "var x=1" not in result
        assert "color:red" not in result
        assert "标题" in result

    def test_read_empty_html(self, temp_document_dir: str) -> None:
        html_path = Path(temp_document_dir) / "empty.html"
        html_path.write_text("<html><body></body></html>", encoding="utf-8")

        reader = HtmlReader()
        result = reader.read(str(html_path))

        assert "No text found" in result

    def test_read_nonexistent_file(self) -> None:
        reader = HtmlReader()
        result = reader.read("/nonexistent/path/page.html")

        assert "Error reading HTML" in result


class TestJsonReader:
    def test_read_json_pretty_printed(self, temp_document_dir: str) -> None:
        json_path = Path(temp_document_dir) / "data.json"
        json_path.write_text('{"name":"张三","age":18}', encoding="utf-8")

        reader = JsonReader()
        result = reader.read(str(json_path))

        assert '"name": "张三"' in result
        assert '"age": 18' in result

    def test_read_invalid_json_returns_raw(self, temp_document_dir: str) -> None:
        json_path = Path(temp_document_dir) / "broken.json"
        json_path.write_text("{not valid json", encoding="utf-8")

        reader = JsonReader()
        result = reader.read(str(json_path))

        assert result == "{not valid json"

    def test_read_empty_json(self, temp_document_dir: str) -> None:
        json_path = Path(temp_document_dir) / "empty.json"
        json_path.write_text("", encoding="utf-8")

        reader = JsonReader()
        result = reader.read(str(json_path))

        assert "No text found" in result

    def test_read_nonexistent_file(self) -> None:
        reader = JsonReader()
        result = reader.read("/nonexistent/path/data.json")

        assert "Error reading JSON" in result


class TestLegacySpreadsheet:
    """旧版电子表格格式（.xls/.xlsb/.ods）测试。"""

    def test_read_real_xls_via_calamine(self, temp_document_dir: str) -> None:
        """测试真实 .xls 文件通过 calamine 读取。"""
        xlwt = pytest.importorskip("xlwt")

        xls_path = Path(temp_document_dir) / "real.xls"
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet("数据")
        sheet.write(0, 0, "姓名")
        sheet.write(0, 1, "年龄")
        sheet.write(1, 0, "张三")
        sheet.write(1, 1, 18)
        workbook.save(str(xls_path))

        reader = ExcelReader()
        result = reader.read(str(xls_path))

        assert "=== Sheet: 数据 ===" in result
        assert "姓名" in result
        assert "张三" in result

    @mock.patch("mcp_documents_reader._extract_text_with_command")
    def test_read_xls_falls_back_to_xls2csv(
        self, mock_command: mock.MagicMock, temp_document_dir: str
    ) -> None:
        """calamine 不可用时回退 xls2csv。"""
        xls_path = Path(temp_document_dir) / "legacy.xls"
        xls_path.write_bytes(b"\xd0\xcf\x11\xe0" + b"\x00" * 64)
        mock_command.return_value = '"姓名","年龄"\n"张三","18"\n'

        with mock.patch.object(ExcelReader, "_read_with_calamine", return_value=None):
            reader = ExcelReader()
            result = reader.read(str(xls_path))

        assert "姓名\t年龄" in result

    def test_read_xls_binary_text_fallback(self, temp_document_dir: str) -> None:
        """测试无任何提取器时回退二进制文本提取。"""
        xls_path = Path(temp_document_dir) / "legacy.xls"
        xls_path.write_bytes(
            b"\xd0\xcf\x11\xe0" + b"\x00" * 64 + "员工名单数据表".encode("utf-16-le")
        )

        reader = ExcelReader()
        with (
            mock.patch.multiple(
                reader,
                _read_with_calamine=lambda *a: None,
                _convert_xls_to_xlsx=lambda *a: None,
            ),
            mock.patch(
                "mcp_documents_reader._extract_text_with_command", return_value=None
            ),
        ):
            result = reader.read(str(xls_path))

        assert "员工名单数据表" in result


class TestRtfReader:
    """RTF 文档读取测试。"""

    def test_read_rtf_basic(self, temp_document_dir: str) -> None:
        rtf_path = Path(temp_document_dir) / "doc.rtf"
        rtf_path.write_text(
            r"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}"
            r"\pard\fs24 Hello RTF \par 第二行内容\par}",
            encoding="utf-8",
        )

        reader = RtfReader()
        result = reader.read(str(rtf_path))

        assert "Hello RTF" in result
        assert "第二行内容" in result
        assert "fonttbl" not in result

    def test_read_rtf_unicode_escapes(self, temp_document_dir: str) -> None:
        rtf_path = Path(temp_document_dir) / "unicode.rtf"
        rtf_path.write_text(
            r"{\rtf1\ansi \uc1\u20013?\u22269?}",  # 中文 and 界
            encoding="utf-8",
        )

        reader = RtfReader()
        result = reader.read(str(rtf_path))

        assert "中" in result or "界" in result

    def test_read_empty_rtf(self, temp_document_dir: str) -> None:
        rtf_path = Path(temp_document_dir) / "empty.rtf"
        rtf_path.write_text(r"{\rtf1\ansi}", encoding="utf-8")

        reader = RtfReader()
        result = reader.read(str(rtf_path))

        assert "No text found" in result

    def test_read_nonexistent_file(self) -> None:
        reader = RtfReader()
        result = reader.read("/nonexistent/path/doc.rtf")

        assert "Error reading RTF" in result


class TestLegacyBinaryExtraction:
    """旧版二进制格式（.doc/.ppt）兜底文本提取测试。"""

    def test_doc_binary_fallback(self, temp_document_dir: str) -> None:
        doc_path = Path(temp_document_dir) / "legacy.doc"
        doc_path.write_bytes(
            b"\xd0\xcf\x11\xe0"
            + b"\x00" * 128
            + "这是旧版文档的正文内容".encode("utf-16-le")
            + b"\x00" * 64
        )

        reader = DocReader()
        with mock.patch.multiple(
            "mcp_documents_reader",
            _extract_text_with_mdls=lambda *a: None,
            _extract_text_with_textutil=lambda *a: None,
            _extract_text_with_command=lambda *a: None,
            _extract_text_with_libreoffice=lambda *a: None,
        ):
            result = reader.read(str(doc_path))

        assert "这是旧版文档的正文内容" in result

    def test_ppt_binary_fallback(self, temp_document_dir: str) -> None:
        ppt_path = Path(temp_document_dir) / "legacy.ppt"
        ppt_path.write_bytes(
            b"\xd0\xcf\x11\xe0"
            + b"\x00" * 128
            + "旧版演示文稿标题".encode("utf-16-le")
            + b"\x00" * 64
        )

        reader = PptReader()
        with mock.patch.multiple(
            "mcp_documents_reader",
            _extract_text_with_mdls=lambda *a: None,
            _extract_text_with_command=lambda *a: None,
            _extract_text_with_libreoffice=lambda *a: None,
        ):
            result = reader.read(str(ppt_path))

        assert "旧版演示文稿标题" in result

    def test_read_rtf_control_syntax(self, temp_document_dir: str) -> None:
        """覆盖 RTF 控制符号、十六进制转义、可忽略目标与特殊控制字。"""
        rtf_path = Path(temp_document_dir) / "syntax.rtf"
        rtf_path.write_text(
            r"{\rtf1\ansi A\'e9B \'zz\{quoted\}\back\~nbsp"
            r"\emdash\endash\lquote\rquote\ldblquote\rdblquote\bullet"
            r"\tab T\line L\row R\page P\sect S"
            r"{\*\shppict {\pict SKIP ME}}\u-1000?\uc0\u20013\par end\}",
            encoding="utf-8",
        )

        reader = RtfReader()
        result = reader.read(str(rtf_path))

        assert "A\u00e9B" in result
        assert "{quoted}" in result
        assert "\u2014" in result
        assert "\u2022" in result
        assert "SKIP ME" not in result
        assert "T\nL" in result or "T\t" not in result or "L" in result
        assert "end" in result

    def test_read_rtf_corrupt_markup(self, temp_document_dir: str) -> None:
        """残缺 RTF（孤立反斜杠、非法控制字）不应抛异常。"""
        rtf_path = Path(temp_document_dir) / "corrupt.rtf"
        rtf_path.write_text(r"{\rtf1 text\!more\\", encoding="utf-8")

        reader = RtfReader()
        result = reader.read(str(rtf_path))

        assert "text" in result


class _FakeOle:
    """Minimal olefile stand-in serving pre-built streams."""

    def __init__(self, streams: dict[str, bytes]) -> None:
        self._streams = streams

    def openstream(self, name: str):
        import io as _io

        if name not in self._streams:
            raise OSError(f"no stream {name}")
        return _io.BytesIO(self._streams[name])

    def close(self) -> None:
        pass


def _build_doc_streams(
    text_utf16: str, compressed_text: bytes = b""
) -> dict[str, bytes]:
    """Build a minimal WordDocument/1Table pair with a valid piece table."""
    import struct as _st

    word_stream = bytearray(0x220)
    # fWhichTblStm flag -> use "1Table"
    word_stream[0x0B] = 0x02
    # ccpText at 0x4C
    _st.pack_into("<I", word_stream, 0x4C, len(text_utf16) + len(compressed_text))

    cps = [0]
    pcds = []
    if text_utf16:
        fc = len(word_stream)
        word_stream += text_utf16.encode("utf-16-le")
        cps.append(cps[-1] + len(text_utf16))
        pcds.append(fc)  # uncompressed -> UTF-16LE
    if compressed_text:
        fc = len(word_stream)
        word_stream += compressed_text
        cps.append(cps[-1] + len(compressed_text))
        pcds.append((fc * 2) | 0x40000000)  # compressed -> cp1252

    pcdt = b"".join(_st.pack("<I", cp) for cp in cps) + b"".join(
        b"\x00\x00" + _st.pack("<I", fc) + b"\x00\x00" for fc in pcds
    )
    clx = b"\x02" + _st.pack("<I", len(pcdt)) + pcdt

    table_stream = bytearray(64)
    _st.pack_into("<I", word_stream, 0x1A2, len(table_stream))  # fcClx
    _st.pack_into("<I", word_stream, 0x1A6, len(clx))  # lcbClx
    table_stream += clx

    return {"WordDocument": bytes(word_stream), "1Table": bytes(table_stream)}


class TestOleExtraction:
    """.doc 分片表与 .ppt 记录流的 OLE 解析测试。"""

    def test_doc_ole_utf16(self, temp_document_dir: str) -> None:
        doc_path = Path(temp_document_dir) / "real.doc"
        doc_path.write_bytes(b"\xd0\xcf\x11\xe0" + b"\x00" * 32)
        streams = _build_doc_streams("第一部分内容\r第二段落文字")

        with (
            mock.patch.multiple(
                "mcp_documents_reader",
                _extract_text_with_mdls=lambda *a: None,
                _extract_text_with_textutil=lambda *a: None,
                _extract_text_with_command=lambda *a: None,
                _extract_text_with_libreoffice=lambda *a: None,
            ),
            mock.patch(
                "mcp_documents_reader._open_ole_streams",
                return_value=_FakeOle(streams),
            ),
        ):
            result = DocReader().read(str(doc_path))

        assert "第一部分内容" in result
        assert "第二段落文字" in result

    def test_doc_ole_compressed_piece(self, temp_document_dir: str) -> None:
        doc_path = Path(temp_document_dir) / "mixed.doc"
        doc_path.write_bytes(b"\xd0\xcf\x11\xe0" + b"\x00" * 32)
        streams = _build_doc_streams("中文标题\r", b"English body text")

        with (
            mock.patch.multiple(
                "mcp_documents_reader",
                _extract_text_with_mdls=lambda *a: None,
                _extract_text_with_textutil=lambda *a: None,
                _extract_text_with_command=lambda *a: None,
                _extract_text_with_libreoffice=lambda *a: None,
            ),
            mock.patch(
                "mcp_documents_reader._open_ole_streams",
                return_value=_FakeOle(streams),
            ),
        ):
            result = DocReader().read(str(doc_path))

        assert "中文标题" in result
        assert "English body text" in result

    def test_ppt_ole_records(self, temp_document_dir: str) -> None:
        import struct as _st

        ppt_path = Path(temp_document_dir) / "real.ppt"
        ppt_path.write_bytes(b"\xd0\xcf\x11\xe0" + b"\x00" * 32)

        def _record(rtype: int, payload: bytes) -> bytes:
            return _st.pack("<HHI", 0, rtype, len(payload)) + payload

        stream = (
            _record(4000, "幻灯片标题".encode("utf-16-le"))
            + _record(4008, b"Bullet point one")
            + _record(4026, "演讲者备注".encode("utf-16-le"))
            + _record(1006, b"\x00" * 10)  # unrelated record type
        )

        with (
            mock.patch.multiple(
                "mcp_documents_reader",
                _extract_text_with_mdls=lambda *a: None,
                _extract_text_with_command=lambda *a: None,
                _extract_text_with_libreoffice=lambda *a: None,
            ),
            mock.patch(
                "mcp_documents_reader._open_ole_streams",
                return_value=_FakeOle({"PowerPoint Document": stream}),
            ),
        ):
            result = PptReader().read(str(ppt_path))

        assert "幻灯片标题" in result
        assert "Bullet point one" in result
        assert "演讲者备注" in result


class TestOdfReader:
    """ODF（.odt/.odp）读取测试。"""

    def _make_odf(self, path: Path) -> Path:
        content = (
            '<?xml version="1.0"?>'
            "<office:document-content "
            'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
            'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">'
            "<office:body><office:text>"
            "<text:p>第一段 ODF 内容</text:p>"
            "<text:p>Second paragraph</text:p>"
            "</office:text></office:body></office:document-content>"
        )
        import zipfile as _zf

        with _zf.ZipFile(path, "w") as z:
            z.writestr("mimetype", "application/vnd.oasis.opendocument.text")
            z.writestr("content.xml", content)
        return path

    def test_read_odt(self, temp_document_dir: str) -> None:
        odt_path = self._make_odf(Path(temp_document_dir) / "doc.odt")
        result = OdfReader().read(str(odt_path))

        assert "第一段 ODF 内容" in result
        assert "Second paragraph" in result

    def test_read_odt_not_zip(self, temp_document_dir: str) -> None:
        bad_path = Path(temp_document_dir) / "bad.odt"
        bad_path.write_bytes(b"not a zip")

        result = OdfReader().read(str(bad_path))

        assert "Error reading ODF" in result

    def test_read_odt_missing_content(self, temp_document_dir: str) -> None:
        import zipfile as _zf

        empty_path = Path(temp_document_dir) / "empty.odt"
        with _zf.ZipFile(empty_path, "w") as z:
            z.writestr("mimetype", "x")

        result = OdfReader().read(str(empty_path))

        assert "Error reading ODF" in result
