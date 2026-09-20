# API 参考

## 目录

- [核心类](#核心类)
  - [DocumentReader](#documentreader)
  - [DocxReader](#docxreader)
  - [PdfReader](#pdfreader)
  - [ExcelReader](#excelreader)
  - [TxtReader](#txtreader)
- [工厂类](#工厂类)
  - [DocumentReaderFactory](#documentreaderfactory)
- [MCP 工具](#mcp-工具)
  - [read_document](#read_document)

---

## 核心类

### DocumentReader

所有文档读取器的抽象基类。

```python
from mcp_documents_reader import DocumentReader

class MyReader(DocumentReader):
    def read(self, file_path: str) -> str:
        # 实现读取逻辑
        return "content"
```

**方法：**

| 方法 | 描述 |
|------|------|
| `read(file_path: str) -> str` | 读取并提取文档文本 |

---

### DocxReader

读取 DOCX（Microsoft Word）文档。

```python
from mcp_documents_reader import DocxReader

reader = DocxReader()
content = reader.read("/path/to/document.docx")
```

**支持扩展名：** `.docx`

**特性：**
- 文本提取
- 表格提取
- 段落格式

---

### PdfReader

读取 PDF 文档。

> **注意：** 从 v1.2.0 开始，PDF 读取器已从 PyPDF2 迁移到 pypdf（更安全、维护更好）。

```python
from mcp_documents_reader import PdfReader

reader = PdfReader()
content = reader.read("/path/to/document.pdf")
```

**支持扩展名：** `.pdf`

**特性：**
- 从 PDF 页面提取文本
- 多页支持

---

### ExcelReader

读取 Excel 电子表格。

```python
from mcp_documents_reader import ExcelReader

reader = ExcelReader()
content = reader.read("/path/to/spreadsheet.xlsx")
```

**支持扩展名：** `.xlsx`, `.xls`, `.xlsm`, `.xlsb`, `.ods`

**特性：**
- 多工作表支持
- 单元格数据提取
- 工作表名称列表

---

### TxtReader

读取纯文本文件，自动检测编码。

```python
from mcp_documents_reader import TxtReader

reader = TxtReader()
content = reader.read("/path/to/file.txt")
```

**支持扩展名：** `.txt`

**特性：**
- 自动编码检测（UTF-8、GBK 等）
- Latin-1 回退处理二进制文件

---

## 工厂类

### DocumentReaderFactory

根据文件扩展名创建适当读取器的工厂类。

```python
from mcp_documents_reader import DocumentReaderFactory

# 获取文件读取器
reader = DocumentReaderFactory.get_reader("document.pdf")

# 检查格式是否支持
is_supported = DocumentReaderFactory.is_supported("document.pdf")

# 获取支持的扩展名列表
readers_map = DocumentReaderFactory._readers
```

**方法：**

| 方法 | 描述 |
|------|------|
| `get_reader(file_path: str) -> DocumentReader` | 获取适合文件的读取器 |
| `is_supported(file_path: str) -> bool` | 检查文件格式是否支持 |

**支持的扩展名：**

| 扩展名 | 读取器类 |
|--------|----------|
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
| `.odt`, `.odp` | OdfReader |
| `.xml`, `.yaml`, `.yml` | TxtReader |

---

## MCP 工具

### read_document

使用统一接口读取任何支持的文档类型。

**参数：**

| 参数 | 类型 | 必需 | 描述 |
|------|------|------|------|
| `filename` | string | 是 | 文档文件路径（绝对路径或相对路径） |

**返回：** 从文档中提取的文本内容。

**示例：**

```python
# 读取 DOCX 文件
content = read_document(filename="report.docx")

# 读取 PDF 文件
content = read_document(filename="paper.pdf")

# 读取 Excel 文件
content = read_document(filename="data.xlsx")

# 读取文本文件
content = read_document(filename="notes.txt")
```

**错误处理：**

- 文件不存在时返回错误信息
- 不支持的格式返回错误信息
- 损坏的文件返回错误信息

### extract_document_images

提取 DOCX 文件中的嵌入图片，并返回结构化 JSON 元数据。

**参数：**

| 参数 | 类型 | 必需 | 描述 |
|------|------|------|------|
| `filename` | string | 是 | DOCX 文件路径 |
| `output_dir` | string | 否 | 图片导出目录（缺省时导出到临时目录） |

**返回：** 包含图片元数据与导出路径的 JSON 字符串。

### read_document_images

提取文档内嵌图片并以图像内容块返回，调用方模型可直接看图识别（无需服务端 OCR）。支持 DOCX / PPTX / XLSX / EPUB / PDF（扫描页即页面图片），以及 `.doc` / `.ppt` / `.xls` 的内嵌图片签名扫描。

**参数：**

| 参数 | 类型 | 必需 | 描述 |
|------|------|------|------|
| `filename` | string | 是 | 文档文件路径 |
| `max_images` | integer | 否 | 最多返回的图片数，默认 20 |

**返回：** 文本摘要后跟每张图片一个图像内容块。

### write_word_document

生成 `.docx` Word 文档，或通过 LibreOffice 转换导出 `.doc`。

**参数：**

| 参数 | 类型 | 必需 | 描述 |
|------|------|------|------|
| `filename` | string | 是 | 输出路径，后缀为 `.docx` 或 `.doc` |
| `title` | string | 否 | 文档标题 |
| `paragraphs` | string[] | 否 | 按顺序写入的段落 |
| `tables` | object[] | 否 | 表格定义，支持 `title`、`headers`、`rows` |

**返回：** 描述生成文件路径与格式的 JSON 字符串。

### write_presentation

生成 `.pptx` 演示文稿，或通过 LibreOffice 转换导出 `.ppt`。

**参数：**

| 参数 | 类型 | 必需 | 描述 |
|------|------|------|------|
| `filename` | string | 是 | 输出路径，后缀为 `.pptx` 或 `.ppt` |
| `title` | string | 否 | 标题页标题 |
| `subtitle` | string | 否 | 标题页副标题 |
| `slides` | object[] | 否 | 幻灯片定义，支持 `title`、`paragraphs`、`bullets`、`table` |

**返回：** 描述生成文件路径与格式的 JSON 字符串。

### write_spreadsheet

生成 `.xlsx` 多工作表电子表格或 `.csv` 文件，或通过 LibreOffice 转换导出 `.xls`。

**参数：**

| 参数 | 类型 | 必需 | 描述 |
|------|------|------|------|
| `filename` | string | 是 | 输出路径，后缀为 `.xlsx`、`.csv` 或 `.xls` |
| `sheets` | object[] | 否 | 工作表定义，支持 `name`、`headers`、`rows` |
| `headers` | string[] | 否 | 单工作表表头（未提供 `sheets` 时生效） |
| `rows` | object[] | 否 | 单工作表数据行（未提供 `sheets` 时生效） |

**返回：** 描述生成文件路径与格式的 JSON 字符串。

**说明：** CSV 输出仅支持单个工作表；数值类型（int/float/bool）会以原生类型写入单元格。

### convert_document

通过 LibreOffice 将文档转换为其他格式。

**参数：**

| 参数 | 类型 | 必需 | 描述 |
|------|------|------|------|
| `filename` | string | 是 | 源文档路径 |
| `target_format` | string | 是 | 目标扩展名，如 `pdf`、`docx`、`txt`、`html`、`csv` |
| `output_dir` | string | 否 | 输出目录，默认为源文件所在目录 |

**返回：** 描述转换后文件路径与格式的 JSON 字符串。

**说明：** 需要安装 LibreOffice 并确保 `soffice` 或 `libreoffice` 在 `PATH` 中。

### list_supported_formats

列出所有支持读取与写入的文件格式。

**参数：** 无。

**返回：** 包含 `read`（可读取扩展名列表）与 `write`（按文档类型分组的可写扩展名）的 JSON 字符串。
