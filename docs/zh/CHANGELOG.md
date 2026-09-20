# 更新日志

本项目的所有重要变更都将记录在此文件中。

格式基于 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.0.0/)，
本项目遵循 [语义化版本](https://semver.org/lang/zh-CN/)。

## [1.5.0] - 2026-09-20

### 新增

- **电子表格生成工具**：新增 `write_spreadsheet`，支持多工作表 `.xlsx`、`.csv`（UTF-8 BOM）生成，以及 `.xls` 的 LibreOffice 转换导出
- **文档转换工具**：新增 `convert_document`，通过 LibreOffice 将文档转换为 `pdf`、`docx`、`txt`、`html`、`csv` 等格式
- **格式清单工具**：新增 `list_supported_formats`，返回可读写格式的 JSON 列表
- **图片直读工具**：新增 `read_document_images`，把文档内嵌图片作为图像内容块直接交给调用方模型识别（覆盖 DOCX / PPTX / XLSX / EPUB / PDF 扫描页，`.doc` / `.ppt` / `.xls` 经签名扫描兜底）
- **新增读取格式**：`.html` / `.htm`（忽略脚本与样式的正文提取）、`.json`（解析并格式化输出）、`.xml` / `.yaml` / `.yml`（文本提取）、`.rtf`（控制字解析）、`.xlsm` / `.xlsb` / `.ods`（电子表格读取）、`.odt` / `.odp`（OpenDocument content.xml 解析）
- **旧版格式结构化解析**：`.doc` 通过 FIB 分片表解析 WordDocument 流、`.ppt` 通过记录流解析文本原子，零依赖即可近乎完整提取正文；最后仍有二进制文本流兜底，保证开箱可读
- **olefile 依赖**：新增 `olefile` 用于 OLE 复合文档流访问
- **calamine 回退**：新增 `python-calamine` 依赖，`.xls` / `.xlsb` / `.ods` 优先经由纯 Python 引擎读取

### 修复

- **`.doc` / `.ppt` 正文识别率低**：原先纯二进制扫描会夹带元数据噪声并丢失顺序，现先经 OLE 结构解析，识别率大幅提升
- **`.xls` 读取失效**：`.xls` 原先错误地走 openpyxl 路径必然报错，现支持 calamine、xls2csv 与 LibreOffice 转换多条回退链路
- **PPTX 幻灯片顺序错误**：原先按文件名数字排序，现改为按 `presentation.xml` 的 `sldIdLst` 顺序读取（缺省时回退文件名排序）
- **DOCX 内容顺序错乱**：段落与表格现按文档实际顺序交错提取，不再将全部表格堆到文末
- **EPUB 路径解析**：manifest `href` 中的 URL 编码与 `#` 片段现会正确解码与剥离
- **BOM 残留**：UTF-8 BOM 不再残留在文本类读取结果中
- **加密 PDF**：对空密码加密的 PDF 会尝试解密后再提取文本
- **`~` 路径**：读取类工具与写入工具一致，支持 `~` 主目录展开

### 变更

- **wheel 打包修复**：构建产物不再把 `tests/`、`docs/` 等仓库文件装入 site-packages，仅发布模块与 `py.typed`
- **类型检查**：修复全部 basedpyright 错误与告警，补充 `py.typed` 标记
- **代码规范**：清理全部 ruff 检查违例并统一格式化

## [1.4.0] - 2026-04-14

### 新增

- **扩展读取能力**：新增 CSV、Markdown、DOC、PPT、PPTX、EPUB 的读取支持
- **Word 生成工具**：新增 `write_word_document`，支持生成带段落和表格的 `.docx`
- **演示文稿生成工具**：新增 `write_presentation`，支持生成带标题页、正文、要点和表格的 `.pptx`
- **DOCX 图片工具公开化**：将 `extract_document_images` 正式纳入对外工具说明

### 变更

- **旧 Office 格式导出链路**：`.doc` 和 `.ppt` 改为基于生成后的 `.docx` / `.pptx` 通过 LibreOffice 转换导出
- **包元信息更新**：同步更新项目描述、关键词、MCP server 元信息和版本号，使其与当前读写能力一致
- **文档刷新**：更新中英文 README 和工具说明，反映当前 MCP 工具面

### 依赖

- 新增 `python-pptx>=0.6.23`，用于演示文稿生成

### 测试

- 新增面向写作生成的 pytest 覆盖，包括 Word 生成、PowerPoint 生成和 LibreOffice 转换回退行为

## [1.3.1] - 2026-03-13

### 安全修复

- **pypdf 安全漏洞**：升级 pypdf>=6.8.0，修复 CVE-2026-28804
  - 修复 ASCIIHexDecode 流解码效率问题，防止 DoS 攻击

### 变更

- **依赖升级**：
  - pypdf>=6.7.4 → pypdf>=6.8.0

## [1.3.0] - 2025-03-10

### 变更

- **灵活的文件路径访问**：移除 `DOCUMENT_DIRECTORY` 限制，现支持绝对路径和相对路径
  - 移除 `DOCUMENT_DIRECTORY` 环境变量依赖
  - 移除 `AppContext` dataclass 和 `app_lifespan` 函数
  - 移除 `_get_document_path()` 安全函数
  - `read_document()` 现直接使用 `Path(filename)` 处理路径
- **简化架构**：移除 FastMCP lifespan 配置，代码更简洁
- **测试套件优化**：删除 `test_lifespan.py`，更新 `test_tools.py` 添加新路径处理测试

### 移除

- `DOCUMENT_DIRECTORY` 环境变量支持
- `AppContext` dataclass
- `app_lifespan` 异步上下文管理器
- `_get_document_path()` 辅助函数

## [1.2.1] - 2025-03-02

### 安全修复

- **pypdf 安全漏洞**：升级 pypdf>=6.7.4，修复 3 个 CVE
  - CVE-2026-28351: RunLengthDecode 流可耗尽 RAM
  - CVE-2026-27888: FlateDecode XFA 流可耗尽 RAM
  - CVE-2026-27628: 循环引用导致无限循环
- **MCP SDK 升级**：升级 mcp>=1.26.0
- **测试代码安全**：重构路径遍历测试代码，避免静态分析误报

### 变更

- **依赖升级**：
  - mcp>=1.23.0 → mcp>=1.26.0
  - pypdf>=6.7.1 → pypdf>=6.7.4
  - typing_extensions>=4.12.0 → typing_extensions>=4.15.0

## [1.2.0] - 2025-03-02

### 安全修复

- **MCP SDK 安全漏洞**：升级 mcp>=1.23.0，修复 3 个高危 CVE
  - CVE-2025-53365: Streamable HTTP Transport 未处理异常导致 DoS
  - CVE-2025-53366: FastMCP Server 验证错误导致 DoS
  - CVE-2025-66416: DNS rebinding 保护默认未启用
- **PyPDF2 安全漏洞**：替换为 pypdf>=6.7.1，修复 CVE-2023-36464
- **路径遍历防护**：添加显式路径验证，防止任意文件读取攻击
- **错误信息脱敏**：移除错误信息中的完整路径，防止信息泄露

### 新增

- **PyPI 包元数据**：添加 project.urls，链接到 GitHub 仓库

### 变更

- **依赖升级**：
  - mcp>=0.1.0 → mcp>=1.23.0
  - PyPDF2>=3.0.1 → pypdf>=6.7.1
  - python-docx>=0.8.11 → python-docx>=1.2.0
  - openpyxl>=3.0.10 → openpyxl>=3.1.5
  - typing_extensions>=4.0.0 → typing_extensions>=4.12.0
- **CI/CD 迁移**：从 pip 迁移到 uv，提升构建速度

## [1.1.0] - 2025-03-01

### 修复

- **Python 兼容性**：使用 `typing_extensions.override` 替代 `typing.override`，兼容 Python 3.10+
- **类型检查**：修复 Basedpyright 类型错误
  - 修复 `openpyxl.Workbook.active` 可选类型检查
  - 修复方法覆写参数名匹配问题
- **编码处理**：移除无效的 `ansi` 编码（Python 标准库不支持）
- **测试修复**：修复路径遍历测试用例

## [1.0.3] - 2025-03-01

### 新增

- **CI/CD 工作流**：添加 GitHub Actions 工作流用于自动化测试和发布
  - CI 工作流：Ruff、Basedpyright、Pytest
  - Release 工作流：发布到 PyPI 和 MCP Registry
  - 支持 Python 3.10-3.14

- **测试套件**：测试覆盖率提升至 95%
  - 102 个测试用例覆盖所有核心模块
  - 所有读取器的单元测试
  - MCP 工具的集成测试

- **文档**：添加完整的文档结构
  - API 参考
  - 用户指南
  - 贡献指南

### 变更

- **类型检查**：切换到 Basedpyright，获得更好的类型推断
- **代码格式化**：使用 Ruff format 替代 Black
- **开发依赖**：更新开发工具链

### 修复

- **类型安全**：修复所有 Basedpyright 类型错误
- **代码质量**：修复所有 Ruff 代码检查问题

## [1.0.2] - 2025-02-28

### 新增

- **MCP 工具**：添加完整的 MCP 工具接口
  - `read_document`：主读取工具
  - 所有文档类型的统一接口

- **错误处理**：改进错误消息和异常处理
  - 不支持格式的更好错误消息
  - 损坏文件的优雅处理

### 变更

- **架构**：使用工厂模式改进读取器架构
- **编码检测**：改进文本文件的自动编码检测

## [1.0.1] - 2025-02-27

### 新增

- **Excel 支持**：添加 Excel 读取器，支持 .xlsx 和 .xls 文件
  - 多工作表支持
  - 单元格数据提取

- **PDF 支持**：添加 PDF 读取器
  - 从 PDF 页面提取文本
  - 多页支持

### 修复

- **编码**：改进文本文件的编码检测
- **错误消息**：更详细的错误消息

## [1.0.0] - 2025-02-25

### 新增

- **首次发布**：MCP Document Reader 首次公开发布
  - 文档读取器抽象基类
  - 使用 python-docx 的 DOCX 读取器
  - 使用 PyPDF2 的 PDF 读取器
  - 使用 openpyxl 的 Excel 读取器
  - 带编码检测的文本读取器
  - 用于读取器选择的工厂模式
  - 支持 AI 助手的 MCP 协议

- **支持的格式**：
  - 输入：DOCX、PDF、Excel（XLSX/XLS）、文本

- **功能特性**：
  - 自动格式检测
  - 文本文件编码检测
  - 损坏文件错误处理
  - AI 助手的 MCP 工具接口
