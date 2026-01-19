# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.2.0] - 2025-01-19

### Added
- **Image Operations** (6 new tools):
  - `get_images_info` - Get metadata for all images in document
  - `read_images` - Read images and return them via MCP Image type for AI viewing
  - `add_image` - Add image at end of document
  - `insert_image_after_paragraph` - Insert image after specific paragraph
  - `delete_image` - Delete image by index
  - `replace_image` - Replace image at specific index

- **Enhanced Table Operations** (5 new tools):
  - `insert_table_after_paragraph` - Insert table after specific paragraph
  - `update_table_cell` - Update cell content with formatting options
  - `delete_table` - Delete entire table by index
  - `add_table_row` - Add row to existing table
  - `delete_table_row` - Delete row from table

- **Docker Support**:
  - Added `Dockerfile` for containerized deployment
  - Added `docker-compose.yml` for easy orchestration
  - Added `.dockerignore` for optimized builds

- **PyPI Packaging**:
  - Added `pyproject.toml` for modern Python packaging
  - Added `main()` entry point for CLI usage

- **Testing**:
  - Added `tests/test_server.py` with comprehensive unit tests

### Changed
- Updated `README.md` with Docker deployment instructions and MCP configuration examples
- Improved error handling across all tools

## [0.1.0] - Initial Release

### Added
- **Document Operations**:
  - `create_new_document` - Create empty Word document
  - `get_document_info` - Get document overview and statistics
  - `read_document_structure` - Read paragraphs with styling info

- **Table Operations**:
  - `read_tables` - Read table content
  - `create_table_with_data` - Create table with initial data

- **Paragraph Operations**:
  - `add_formatted_paragraph` - Add paragraph with formatting
  - `replace_paragraph` - Replace paragraph content
  - `insert_paragraph_after` - Insert paragraph at position
  - `search_and_replace` - Find and replace text

---

## Tool Count Summary

| Version | Total Tools |
|---------|-------------|
| 0.1.0   | 9           |
| 0.2.0   | 20          |
