# API Reference

## Table of Contents

- [Core Classes](#core-classes)
  - [DocumentReader](#documentreader)
  - [DocxReader](#docxreader)
  - [PdfReader](#pdfreader)
  - [ExcelReader](#excelreader)
  - [TxtReader](#txtreader)
- [Factory Class](#factory-class)
  - [DocumentReaderFactory](#documentreaderfactory)
- [MCP Tools](#mcp-tools)
  - [read_document](#read_document)

---

## Core Classes

### DocumentReader

Abstract base class for all document readers.

```python
from mcp_documents_reader import DocumentReader

class MyReader(DocumentReader):
    def read(self, file_path: str) -> str:
        # Implement reading logic
        return "content"
```

**Methods:**

| Method | Description |
|--------|-------------|
| `read(file_path: str) -> str` | Read and extract text from the document |

---

### DocxReader

Reads DOCX (Microsoft Word) documents.

```python
from mcp_documents_reader import DocxReader

reader = DocxReader()
content = reader.read("/path/to/document.docx")
```

**Supported Extensions:** `.docx`

**Features:**
- Text extraction
- Table extraction
- Paragraph formatting

---

### PdfReader

Reads PDF documents.

> **Note:** Starting from v1.2.0, the PDF reader has been migrated from PyPDF2 to pypdf (more secure and better maintained).

```python
from mcp_documents_reader import PdfReader

reader = PdfReader()
content = reader.read("/path/to/document.pdf")
```

**Supported Extensions:** `.pdf`

**Features:**
- Text extraction from PDF pages
- Multi-page support

---

### ExcelReader

Reads Excel spreadsheets.

```python
from mcp_documents_reader import ExcelReader

reader = ExcelReader()
content = reader.read("/path/to/spreadsheet.xlsx")
```

**Supported Extensions:** `.xlsx`, `.xls`, `.xlsm`, `.xlsb`, `.ods`

**Features:**
- Multi-sheet support
- Cell data extraction
- Sheet name listing

---

### TxtReader

Reads plain text files with automatic encoding detection.

```python
from mcp_documents_reader import TxtReader

reader = TxtReader()
content = reader.read("/path/to/file.txt")
```

**Supported Extensions:** `.txt`

**Features:**
- Automatic encoding detection (UTF-8, GBK, etc.)
- Latin-1 fallback for binary files

---

## Factory Class

### DocumentReaderFactory

Factory class for creating appropriate readers based on file extension.

```python
from mcp_documents_reader import DocumentReaderFactory

# Get reader for a file
reader = DocumentReaderFactory.get_reader("document.pdf")

# Check if format is supported
is_supported = DocumentReaderFactory.is_supported("document.pdf")

# Get list of supported extensions
readers_map = DocumentReaderFactory._readers
```

**Methods:**

| Method | Description |
|--------|-------------|
| `get_reader(file_path: str) -> DocumentReader` | Get appropriate reader for the file |
| `is_supported(file_path: str) -> bool` | Check if the file format is supported |

**Supported Extensions:**

| Extension | Reader Class |
|-----------|--------------|
| `.txt` | TxtReader |
| `.csv` | CsvReader |
| `.md`, `.markdown` | MarkdownReader |
| `.doc` | DocReader |
| `.docx` | DocxReader |
| `.pdf` | PdfReader |
| `.ppt` | PptReader |
| `.pptx` | PptxReader |
| `.epub` | EpubReader |
| `.xlsx` | ExcelReader |
| `.xls` | ExcelReader |
| `.xlsm` | ExcelReader |
| `.xlsb` | ExcelReader |
| `.ods` | ExcelReader |
| `.html`, `.htm` | HtmlReader |
| `.json` | JsonReader |
| `.rtf` | RtfReader |
| `.xml`, `.yaml`, `.yml` | TxtReader |

---

## MCP Tools

### read_document

Read any supported document type with a unified interface.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `filename` | string | Yes | Document file path (absolute or relative) |

**Returns:** Extracted text content from the document.

**Example:**

```python
# Read a DOCX file
content = read_document(filename="report.docx")

# Read a PDF file
content = read_document(filename="paper.pdf")

# Read an Excel file
content = read_document(filename="data.xlsx")

# Read a text file
content = read_document(filename="notes.txt")
```

**Error Handling:**

- Returns error message if file not found
- Returns error message for unsupported formats
- Returns error message for corrupted files

### extract_document_images

Extract embedded images from a DOCX file and return structured JSON metadata.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `filename` | string | Yes | Path to the DOCX file |
| `output_dir` | string | No | Directory for exported images (defaults to a temp directory) |

**Returns:** JSON string with image metadata and exported file paths.

### write_word_document

Generate a `.docx` Word document, or export `.doc` via LibreOffice conversion.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `filename` | string | Yes | Output path ending with `.docx` or `.doc` |
| `title` | string | No | Document title |
| `paragraphs` | string[] | No | Paragraphs written in order |
| `tables` | object[] | No | Table specs with `title`, `headers`, and `rows` |

**Returns:** JSON string describing the generated file path and format.

### write_presentation

Generate a `.pptx` presentation, or export `.ppt` via LibreOffice conversion.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `filename` | string | Yes | Output path ending with `.pptx` or `.ppt` |
| `title` | string | No | Title slide title |
| `subtitle` | string | No | Title slide subtitle |
| `slides` | object[] | No | Slide specs with `title`, `paragraphs`, `bullets`, `table` |

**Returns:** JSON string describing the generated file path and format.

### write_spreadsheet

Generate a multi-sheet `.xlsx` spreadsheet or `.csv` file, or export `.xls` via LibreOffice conversion.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `filename` | string | Yes | Output path ending with `.xlsx`, `.csv`, or `.xls` |
| `sheets` | object[] | No | Sheet specs with `name`, `headers`, and `rows` |
| `headers` | string[] | No | Single-sheet headers (used when `sheets` is empty) |
| `rows` | object[] | No | Single-sheet data rows (used when `sheets` is empty) |

**Returns:** JSON string describing the generated file path and format.

**Notes:** CSV output supports a single sheet only; numeric values (int/float/bool) are written as native cell types.

### convert_document

Convert a document to another format via LibreOffice.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `filename` | string | Yes | Source document path |
| `target_format` | string | Yes | Target extension such as `pdf`, `docx`, `txt`, `html`, `csv` |
| `output_dir` | string | No | Output directory (defaults to the source directory) |

**Returns:** JSON string describing the converted file path and format.

**Notes:** Requires LibreOffice with `soffice` or `libreoffice` in `PATH`.

### list_supported_formats

List all document formats supported for reading and writing.

**Parameters:** None.

**Returns:** JSON string with `read` (readable extensions) and `write` (writable extensions grouped by document type).
