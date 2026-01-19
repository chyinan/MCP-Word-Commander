# MCP Word Commander

强大的 Word 文档处理 MCP 服务器，支持图片读取/插入、表格操作等功能。

## 简介

MCP Word Commander 是一个基于 MCP (Model Context Protocol) 的 Word 文档处理服务，允许 AI 助手直接读取、编辑 Word 文档，包括：

- **图片处理**：读取图片并直接展示给 AI、插入图片、删除/替换图片
- **表格操作**：读取表格、在指定位置插入表格、修改单元格、添加/删除行
- **段落编辑**：添加、替换、插入段落，支持完整的格式设置
- **搜索替换**：在文档中搜索并替换文本

## 依赖

- Python 3.10+
- 见 `requirements.txt`

## 快速开始

```bash
# Windows
.venv\Scripts\activate
pip install -r requirements.txt
python server.py

# macOS / Linux
source venv/bin/activate
pip install -r requirements.txt
python server.py
```

## 功能列表

### 文档基础操作

| 工具 | 功能 |
|------|------|
| `create_new_document` | 创建新的空白 Word 文档 |
| `get_document_info` | 获取文档基本信息（段落数、表格数等） |
| `read_document_structure` | 分段读取文档内容和样式 |

### 段落操作

| 工具 | 功能 |
|------|------|
| `add_formatted_paragraph` | 在文档末尾添加带格式的段落 |
| `replace_paragraph` | 替换指定段落的内容 |
| `insert_paragraph_after` | 在指定段落后插入新段落 |
| `search_and_replace` | 搜索并替换文本 |

### 图片操作

| 工具 | 功能 |
|------|------|
| `get_images_info` | 获取文档中所有图片的元信息 |
| `read_images` | **读取图片并直接返回给 AI 查看** |
| `add_image` | 在文档末尾添加图片 |
| `insert_image_after_paragraph` | 在指定段落后插入图片 |
| `delete_image` | 删除指定索引的图片 |
| `replace_image` | 替换指定索引的图片 |

### 表格操作

| 工具 | 功能 |
|------|------|
| `read_tables` | 读取表格内容 |
| `create_table_with_data` | 在文档末尾创建表格 |
| `insert_table_after_paragraph` | 在指定段落后插入表格 |
| `update_table_cell` | 修改表格单元格内容 |
| `add_table_row` | 向表格添加新行 |
| `delete_table_row` | 删除表格中的行 |
| `delete_table` | 删除整个表格 |

## 使用示例

### 读取文档中的图片

```python
# AI 可以直接"看到"文档中的图片
read_images("document.docx")

# 只读取第一张图片
read_images("document.docx", image_index=0)
```

### 在指定位置插入图片

```python
# 在第 3 段后插入图片，设置宽度为 4 英寸，居中对齐
insert_image_after_paragraph(
    file_path="document.docx",
    after_index=2,
    image_path="image.png",
    width_inches=4.0,
    alignment="CENTER"
)
```

### 在指定位置插入表格

```python
# 在第 5 段后插入 3x3 表格
insert_table_after_paragraph(
    file_path="document.docx",
    after_index=4,
    rows=3,
    cols=3,
    data=[
        ["姓名", "年龄", "城市"],
        ["张三", "25", "北京"],
        ["李四", "30", "上海"]
    ],
    header_bold=True
)
```

### 修改表格单元格

```python
# 修改第一个表格的 (1, 2) 单元格
update_table_cell(
    file_path="document.docx",
    table_index=0,
    row=1,
    col=2,
    new_text="新内容",
    font_name="Microsoft YaHei",
    font_size=12,
    is_bold=True
)
```

## 项目结构

```
mcp_word_edit/
├── server.py           # MCP 服务器主文件
├── requirements.txt    # Python 依赖
├── README.md          # 说明文档
└── .venv/             # 虚拟环境
```

## 技术特点

- **图片直接展示**：使用 MCP 的 `Image` 类型，AI 可以直接"看到"文档中的图片内容
- **支持中文字体**：完整支持中文字体设置（宋体、黑体等）
- **灵活的位置插入**：支持在任意段落后插入图片和表格
- **完整的表格操作**：支持增删改查表格及其内容

## 贡献

欢迎开 issue 或 PR。提交前请确保代码风格一致并包含必要说明。

## 许可证

本项目使用 MIT 许可证，详见 `LICENSE`。

---

# MCP Word Commander (English)

A powerful Word document processing MCP server with image and table support.

## Overview

MCP Word Commander is a Word document processing service based on MCP (Model Context Protocol), enabling AI assistants to directly read and edit Word documents, including:

- **Image Processing**: Read images and display them directly to AI, insert/delete/replace images
- **Table Operations**: Read tables, insert tables at specific positions, modify cells, add/delete rows
- **Paragraph Editing**: Add, replace, insert paragraphs with full formatting support
- **Search & Replace**: Find and replace text in documents

## Requirements

- Python 3.10+
- See `requirements.txt`

## Quick Start

```bash
# Windows
.venv\Scripts\activate
pip install -r requirements.txt
python server.py

# macOS / Linux
source venv/bin/activate
pip install -r requirements.txt
python server.py
```

## Feature List

### Basic Document Operations

| Tool | Function |
|------|----------|
| `create_new_document` | Create a new blank Word document |
| `get_document_info` | Get basic document information (paragraph count, table count, etc.) |
| `read_document_structure` | Read document content and styles in segments |

### Paragraph Operations

| Tool | Function |
|------|----------|
| `add_formatted_paragraph` | Add a formatted paragraph at the end of document |
| `replace_paragraph` | Replace content of a specific paragraph |
| `insert_paragraph_after` | Insert a new paragraph after a specific paragraph |
| `search_and_replace` | Search and replace text |

### Image Operations

| Tool | Function |
|------|----------|
| `get_images_info` | Get metadata of all images in the document |
| `read_images` | **Read images and return them directly for AI to view** |
| `add_image` | Add image at the end of document |
| `insert_image_after_paragraph` | Insert image after a specific paragraph |
| `delete_image` | Delete image by index |
| `replace_image` | Replace image by index |

### Table Operations

| Tool | Function |
|------|----------|
| `read_tables` | Read table content |
| `create_table_with_data` | Create table at the end of document |
| `insert_table_after_paragraph` | Insert table after a specific paragraph |
| `update_table_cell` | Modify table cell content |
| `add_table_row` | Add new row to table |
| `delete_table_row` | Delete row from table |
| `delete_table` | Delete entire table |

## Usage Examples

### Read Images from Document

```python
# AI can directly "see" images in the document
read_images("document.docx")

# Read only the first image
read_images("document.docx", image_index=0)
```

### Insert Image at Specific Position

```python
# Insert image after paragraph 3, set width to 4 inches, center alignment
insert_image_after_paragraph(
    file_path="document.docx",
    after_index=2,
    image_path="image.png",
    width_inches=4.0,
    alignment="CENTER"
)
```

### Insert Table at Specific Position

```python
# Insert a 3x3 table after paragraph 5
insert_table_after_paragraph(
    file_path="document.docx",
    after_index=4,
    rows=3,
    cols=3,
    data=[
        ["Name", "Age", "City"],
        ["John", "25", "Beijing"],
        ["Jane", "30", "Shanghai"]
    ],
    header_bold=True
)
```

### Modify Table Cell

```python
# Modify cell (1, 2) of the first table
update_table_cell(
    file_path="document.docx",
    table_index=0,
    row=1,
    col=2,
    new_text="New Content",
    font_name="Microsoft YaHei",
    font_size=12,
    is_bold=True
)
```

## Project Structure

```
mcp_word_edit/
├── server.py           # MCP server main file
├── requirements.txt    # Python dependencies
├── README.md          # Documentation
└── .venv/             # Virtual environment
```

## Technical Features

- **Direct Image Display**: Uses MCP's `Image` type, allowing AI to directly "see" image content in documents
- **Chinese Font Support**: Full support for Chinese font settings (SimSun, SimHei, etc.)
- **Flexible Position Insertion**: Support inserting images and tables after any paragraph
- **Complete Table Operations**: Support CRUD operations for tables and their content

## Contributing

Issues and PRs are welcome. Please ensure code style consistency and include necessary documentation before submitting.

## License

This project is licensed under the MIT License. See `LICENSE` for details.
