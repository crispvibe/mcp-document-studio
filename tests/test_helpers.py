from pathlib import Path
from types import SimpleNamespace
from unittest import mock
import zipfile

import pytest

import mcp_documents_reader as reader_module


def test_format_docx_content_with_images_without_images() -> None:
    result = reader_module._format_docx_content_with_images("plain", {"images": []})

    assert result == "plain"


def test_format_docx_content_with_images_without_text() -> None:
    payload = {
        "output_dir": "/tmp/images",
        "images": [
            {
                "index": 1,
                "saved_filename": "image.png",
                "saved_path": "/tmp/images/image.png",
                "saved_uri": "file:///tmp/images/image.png",
                "width_px": 10,
                "height_px": 20,
            },
            "invalid",
        ],
    }

    result = reader_module._format_docx_content_with_images("", payload)

    assert "=== Extracted Images ===" in result
    assert "Temporary image directory: /tmp/images" in result
    assert "image.png" in result


def test_read_text_file_returns_decode_error_when_all_encodings_fail() -> None:
    mocked_open = mock.MagicMock(side_effect=[UnicodeDecodeError("utf-8", b"", 0, 1, "boom")] * len(reader_module.TEXT_ENCODINGS))

    with mock.patch("builtins.open", mocked_open):
        result = reader_module._read_text_file(
            "bad.txt",
            empty_message="EMPTY",
            error_prefix="ERR",
        )

    assert result == "ERR: Could not decode file with any supported encoding."


def test_read_text_file_returns_runtime_error() -> None:
    with mock.patch("builtins.open", side_effect=OSError("denied")):
        result = reader_module._read_text_file(
            "bad.txt",
            empty_message="EMPTY",
            error_prefix="ERR",
        )

    assert result == "ERR: denied"


def test_normalize_and_extract_markup_helpers() -> None:
    assert reader_module._normalize_text_chunks(["  ", " a ", "b  "], "EMPTY") == "a\nb"
    assert reader_module._normalize_text_chunks(["", "   "], "EMPTY") == "EMPTY"
    assert reader_module._local_name("{urn:test}tag") == "tag"
    assert reader_module._extract_markup_text(b"<root><item> first text </item><item>second</item></root>") == "first text\nsecond"


@mock.patch("mcp_documents_reader.subprocess.run")
def test_run_text_command_variants(mock_run: mock.MagicMock) -> None:
    mock_run.return_value = SimpleNamespace(returncode=0, stdout="hello\n")
    assert reader_module._run_text_command(["cmd"]) == "hello"

    mock_run.return_value = SimpleNamespace(returncode=1, stdout="hello\n")
    assert reader_module._run_text_command(["cmd"]) is None

    mock_run.return_value = SimpleNamespace(returncode=0, stdout="(null)\n")
    assert reader_module._run_text_command(["cmd"]) is None

    mock_run.side_effect = RuntimeError("boom")
    assert reader_module._run_text_command(["cmd"]) is None


def test_mdls_and_text_commands(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(reader_module.sys, "platform", "linux")
    assert reader_module._extract_text_with_mdls("file.doc") is None

    monkeypatch.setattr(reader_module.sys, "platform", "darwin")
    monkeypatch.setattr(reader_module.shutil, "which", lambda name: "/usr/bin/mdls" if name == "mdls" else None)
    monkeypatch.setattr(reader_module, "_run_text_command", lambda command: "mdls text")
    assert reader_module._extract_text_with_mdls("file.doc") == "mdls text"

    monkeypatch.setattr(reader_module.shutil, "which", lambda name: None)
    assert reader_module._extract_text_with_textutil("file.doc") is None
    assert reader_module._extract_text_with_command("file.ppt", "catppt") is None

    monkeypatch.setattr(reader_module.shutil, "which", lambda name: f"/usr/bin/{name}")
    monkeypatch.setattr(reader_module, "_run_text_command", lambda command: "command text")
    assert reader_module._extract_text_with_textutil("file.doc") == "command text"
    assert reader_module._extract_text_with_command("file.ppt", "catppt") == "command text"


def test_extract_text_with_libreoffice(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setattr(reader_module.shutil, "which", lambda name: None)
    assert reader_module._extract_text_with_libreoffice("sample.doc") is None

    monkeypatch.setattr(reader_module.shutil, "which", lambda name: "/usr/bin/soffice")

    def fake_run_missing(*args, **kwargs):
        return SimpleNamespace(returncode=0, stdout="", stderr="")

    monkeypatch.setattr(reader_module.subprocess, "run", fake_run_missing)
    assert reader_module._extract_text_with_libreoffice("sample.doc") is None

    def fake_run_success(command, **kwargs):
        outdir = Path(command[command.index("--outdir") + 1])
        (outdir / "sample.txt").write_text("converted text", encoding="utf-8")
        return SimpleNamespace(returncode=0, stdout="", stderr="")

    monkeypatch.setattr(reader_module.subprocess, "run", fake_run_success)
    assert reader_module._extract_text_with_libreoffice("sample.doc") == "converted text"


def test_extract_epub_document_paths_and_fallback(tmp_path: Path) -> None:
    fallback_epub = tmp_path / "fallback.epub"
    with zipfile.ZipFile(fallback_epub, "w") as archive:
        archive.writestr(
            "META-INF/container.xml",
            "<?xml version='1.0'?><container xmlns='urn:oasis:names:tc:opendocument:xmlns:container'><rootfiles><rootfile full-path='OPS/content.opf' media-type='application/oebps-package+xml'/></rootfiles></container>",
        )
        archive.writestr(
            "OPS/content.opf",
            "<?xml version='1.0'?><package xmlns='http://www.idpf.org/2007/opf'><manifest></manifest><spine></spine></package>",
        )
        archive.writestr("OPS/page.xhtml", "<html><body><p>fallback</p></body></html>")

    with zipfile.ZipFile(fallback_epub) as archive:
        assert reader_module._extract_epub_document_paths(archive) == ["OPS/page.xhtml"]

    broken_epub = tmp_path / "broken.epub"
    with zipfile.ZipFile(broken_epub, "w") as archive:
        archive.writestr(
            "META-INF/container.xml",
            "<?xml version='1.0'?><container xmlns='urn:oasis:names:tc:opendocument:xmlns:container'><rootfiles></rootfiles></container>",
        )

    with zipfile.ZipFile(broken_epub) as archive:
        with pytest.raises(ValueError, match="EPUB rootfile not found"):
            reader_module._extract_epub_document_paths(archive)


def test_reader_error_branches(tmp_path: Path) -> None:
    invalid_pptx = tmp_path / "broken.pptx"
    invalid_pptx.write_text("not-a-zip", encoding="utf-8")
    assert "Error reading PPTX" in reader_module.PptxReader().read(str(invalid_pptx))

    invalid_epub = tmp_path / "broken.epub"
    invalid_epub.write_text("not-a-zip", encoding="utf-8")
    assert "Error reading EPUB" in reader_module.EpubReader().read(str(invalid_epub))


def test_doc_and_ppt_reader_error_branches() -> None:
    with mock.patch("mcp_documents_reader._extract_text_with_mdls", return_value=None), mock.patch(
        "mcp_documents_reader._extract_text_with_textutil", return_value=None
    ), mock.patch("mcp_documents_reader._extract_text_with_command", return_value=None), mock.patch(
        "mcp_documents_reader._extract_text_with_libreoffice", return_value=None
    ):
        assert "No available extractor succeeded" in reader_module.DocReader().read("legacy.doc")
        assert "No available extractor succeeded" in reader_module.PptReader().read("slides.ppt")

    with mock.patch("mcp_documents_reader._extract_text_with_mdls", side_effect=RuntimeError("boom")):
        assert "Error reading DOC: boom" == reader_module.DocReader().read("legacy.doc")
        assert "Error reading PPT: boom" == reader_module.PptReader().read("slides.ppt")


def test_extract_document_images_error_and_main(sample_docx_file: Path) -> None:
    with mock.patch.object(reader_module.DocxReader, "extract_images", side_effect=RuntimeError("boom")):
        result = reader_module.extract_document_images(str(sample_docx_file))

    assert result == "Error extracting document images: boom"

    with mock.patch.object(reader_module.mcp, "run") as mock_run:
        reader_module.main()

    mock_run.assert_called_once()
