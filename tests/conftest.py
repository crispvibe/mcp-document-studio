"""pytest 配置文件和共享 fixtures。

本模块提供测试所需的共享配置和 fixture 对象，包括：
- 测试文档文件路径
- 临时目录管理
"""

import tempfile
import zipfile
from pathlib import Path
from textwrap import dedent
from typing import Generator

import pytest

PROJECT_ROOT = Path(__file__).parent.parent
FIXTURES_DIR = PROJECT_ROOT / "tests" / "fixtures"


def _create_sample_pptx(target_path: Path) -> Path:
    with zipfile.ZipFile(target_path, "w") as archive:
        archive.writestr(
            "[Content_Types].xml",
            dedent(
                """\
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """
            ).strip(),
        )
        archive.writestr(
            "_rels/.rels",
            dedent(
                """\
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
                </Relationships>
                """
            ).strip(),
        )
        archive.writestr(
            "ppt/presentation.xml",
            dedent(
                """\
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:sldIdLst>
                    <p:sldId id="256" r:id="rId1"/>
                  </p:sldIdLst>
                </p:presentation>
                """
            ).strip(),
        )
        archive.writestr(
            "ppt/_rels/presentation.xml.rels",
            dedent(
                """\
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                </Relationships>
                """
            ).strip(),
        )
        archive.writestr(
            "ppt/slides/slide1.xml",
            dedent(
                """\
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:sp>
                        <p:txBody>
                          <a:p>
                            <a:r><a:t>演示文稿标题</a:t></a:r>
                          </a:p>
                          <a:p>
                            <a:r><a:t>第二行内容</a:t></a:r>
                          </a:p>
                        </p:txBody>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
            ).strip(),
        )
    return target_path


def _create_sample_epub(target_path: Path) -> Path:
    with zipfile.ZipFile(target_path, "w") as archive:
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
                <package version="2.0" xmlns="http://www.idpf.org/2007/opf" unique-identifier="BookId">
                  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
                    <dc:title>示例 EPUB</dc:title>
                    <dc:identifier id="BookId">sample-book</dc:identifier>
                  </metadata>
                  <manifest>
                    <item id="chapter1" href="chapter1.xhtml" media-type="application/xhtml+xml"/>
                  </manifest>
                  <spine>
                    <itemref idref="chapter1"/>
                  </spine>
                </package>
                """
            ).strip(),
        )
        archive.writestr(
            "OEBPS/chapter1.xhtml",
            dedent(
                """\
                <?xml version="1.0" encoding="UTF-8"?>
                <html xmlns="http://www.w3.org/1999/xhtml">
                  <head><title>第一章</title></head>
                  <body>
                    <h1>第一章标题</h1>
                    <p>这是 EPUB 正文。</p>
                  </body>
                </html>
                """
            ).strip(),
        )
    return target_path


@pytest.fixture
def fixtures_dir() -> Path:
    """获取测试 fixtures 目录路径。

    Returns:
        Path: fixtures 目录的 Path 对象
    """
    return FIXTURES_DIR


@pytest.fixture
def sample_txt_file(fixtures_dir: Path) -> Path:
    """获取示例 TXT 文件路径。

    Args:
        fixtures_dir: fixtures 目录路径

    Returns:
        Path: 示例 TXT 文件路径
    """
    return fixtures_dir / "sample.txt"


@pytest.fixture
def sample_txt_gbk_file(fixtures_dir: Path) -> Path:
    """获取 GBK 编码的示例 TXT 文件路径。

    Args:
        fixtures_dir: fixtures 目录路径

    Returns:
        Path: GBK 编码的示例 TXT 文件路径
    """
    return fixtures_dir / "sample_gbk.txt"


@pytest.fixture
def sample_docx_file(fixtures_dir: Path) -> Path:
    """获取示例 DOCX 文件路径。

    Args:
        fixtures_dir: fixtures 目录路径

    Returns:
        Path: 示例 DOCX 文件路径
    """
    return fixtures_dir / "sample.docx"


@pytest.fixture
def sample_pdf_file(fixtures_dir: Path) -> Path:
    """获取示例 PDF 文件路径。

    Args:
        fixtures_dir: fixtures 目录路径

    Returns:
        Path: 示例 PDF 文件路径
    """
    return fixtures_dir / "sample.pdf"


@pytest.fixture
def sample_excel_file(fixtures_dir: Path) -> Path:
    """获取示例 Excel 文件路径。

    Args:
        fixtures_dir: fixtures 目录路径

    Returns:
        Path: 示例 Excel 文件路径
    """
    return fixtures_dir / "sample.xlsx"


@pytest.fixture
def empty_txt_file(fixtures_dir: Path) -> Path:
    """获取空 TXT 文件路径。

    Args:
        fixtures_dir: fixtures 目录路径

    Returns:
        Path: 空 TXT 文件路径
    """
    return fixtures_dir / "empty.txt"


@pytest.fixture
def empty_docx_file(fixtures_dir: Path) -> Path:
    """获取空 DOCX 文件路径。

    Args:
        fixtures_dir: fixtures 目录路径

    Returns:
        Path: 空 DOCX 文件路径
    """
    return fixtures_dir / "empty.docx"


@pytest.fixture
def empty_pdf_file(fixtures_dir: Path) -> Path:
    """获取空 PDF 文件路径。

    Args:
        fixtures_dir: fixtures 目录路径

    Returns:
        Path: 空 PDF 文件路径
    """
    return fixtures_dir / "empty.pdf"


@pytest.fixture
def empty_excel_file(fixtures_dir: Path) -> Path:
    """获取空 Excel 文件路径。

    Args:
        fixtures_dir: fixtures 目录路径

    Returns:
        Path: 空 Excel 文件路径
    """
    return fixtures_dir / "empty.xlsx"


@pytest.fixture
def temp_document_dir() -> Generator[str, None, None]:
    """创建临时文档目录。

    Yields:
        str: 临时目录路径
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        yield tmpdir


@pytest.fixture
def sample_csv_file(temp_document_dir: str) -> Path:
    csv_path = Path(temp_document_dir) / "sample.csv"
    csv_path.write_text("姓名,年龄\n张三,18\n李四,20\n", encoding="utf-8")
    return csv_path


@pytest.fixture
def sample_markdown_file(temp_document_dir: str) -> Path:
    markdown_path = Path(temp_document_dir) / "sample.md"
    markdown_path.write_text("# 标题\n\n- 第一项\n- 第二项\n", encoding="utf-8")
    return markdown_path


@pytest.fixture
def sample_pptx_file(temp_document_dir: str) -> Path:
    return _create_sample_pptx(Path(temp_document_dir) / "sample.pptx")


@pytest.fixture
def sample_epub_file(temp_document_dir: str) -> Path:
    return _create_sample_epub(Path(temp_document_dir) / "sample.epub")
