"""
Unit tests for mcp-word-edit server.

Tests cover basic document operations including:
- Document creation
- Document info retrieval
- Document structure reading
- Table operations
- Paragraph operations
- Image operations
"""

import os
import sys
import json
import tempfile
import pytest
from pathlib import Path

# Add parent directory to path for import
sys.path.insert(0, str(Path(__file__).parent.parent))

from server import (
    create_new_document,
    get_document_info,
    read_document_structure,
    read_tables,
    add_formatted_paragraph,
    replace_paragraph,
    insert_paragraph_after,
    search_and_replace,
    create_table_with_data,
    insert_table_after_paragraph,
    update_table_cell,
    delete_table,
    add_table_row,
    delete_table_row,
    get_images_info,
)


class TestDocumentCreation:
    """Tests for document creation functionality."""

    def test_create_new_document(self):
        """Test creating a new empty document."""
        with tempfile.TemporaryDirectory() as tmpdir:
            file_path = os.path.join(tmpdir, "test.docx")
            result = create_new_document(file_path)
            
            assert "Successfully created" in result
            assert os.path.exists(file_path)

    def test_create_document_in_nonexistent_dir(self):
        """Test creating document fails gracefully for invalid path."""
        result = create_new_document("/nonexistent/path/test.docx")
        assert "Error" in result or "error" in result.lower()


class TestDocumentInfo:
    """Tests for document info retrieval."""

    def test_get_document_info(self):
        """Test getting document information."""
        with tempfile.TemporaryDirectory() as tmpdir:
            file_path = os.path.join(tmpdir, "test.docx")
            create_new_document(file_path)
            
            result = get_document_info(file_path)
            info = json.loads(result)
            
            assert "total_paragraphs" in info
            assert "total_tables" in info
            assert "file_path" in info

    def test_get_info_nonexistent_file(self):
        """Test getting info for nonexistent file."""
        result = get_document_info("/nonexistent/test.docx")
        info = json.loads(result)
        
        assert "error" in info


class TestParagraphOperations:
    """Tests for paragraph-related operations."""

    def test_add_formatted_paragraph(self):
        """Test adding a formatted paragraph."""
        with tempfile.TemporaryDirectory() as tmpdir:
            file_path = os.path.join(tmpdir, "test.docx")
            create_new_document(file_path)
            
            result = add_formatted_paragraph(
                file_path=file_path,
                text="Hello World",
                font_name="SimSun",
                font_size=12.0,
                is_bold=True,
                alignment="CENTER"
            )
            result_data = json.loads(result)
            
            assert result_data.get("success") == True

    def test_read_document_structure(self):
        """Test reading document structure."""
        with tempfile.TemporaryDirectory() as tmpdir:
            file_path = os.path.join(tmpdir, "test.docx")
            create_new_document(file_path)
            add_formatted_paragraph(file_path, "Test paragraph 1")
            add_formatted_paragraph(file_path, "Test paragraph 2")
            
            result = read_document_structure(file_path)
            data = json.loads(result)
            
            assert "total_paragraphs" in data
            assert "paragraphs" in data

    def test_insert_paragraph_after(self):
        """Test inserting paragraph after specific index."""
        with tempfile.TemporaryDirectory() as tmpdir:
            file_path = os.path.join(tmpdir, "test.docx")
            create_new_document(file_path)
            add_formatted_paragraph(file_path, "First paragraph")
            add_formatted_paragraph(file_path, "Third paragraph")
            
            result = insert_paragraph_after(
                file_path=file_path,
                after_index=0,
                text="Second paragraph (inserted)"
            )
            result_data = json.loads(result)
            
            assert result_data.get("success") == True

    def test_search_and_replace(self):
        """Test search and replace functionality."""
        with tempfile.TemporaryDirectory() as tmpdir:
            file_path = os.path.join(tmpdir, "test.docx")
            create_new_document(file_path)
            add_formatted_paragraph(file_path, "Hello World, Hello Everyone")
            
            result = search_and_replace(
                file_path=file_path,
                search_text="Hello",
                replace_text="Hi"
            )
            result_data = json.loads(result)
            
            assert result_data.get("success") == True
            assert result_data.get("replacements_made", 0) > 0


class TestTableOperations:
    """Tests for table-related operations."""

    def test_create_table_with_data(self):
        """Test creating a table with data."""
        with tempfile.TemporaryDirectory() as tmpdir:
            file_path = os.path.join(tmpdir, "test.docx")
            create_new_document(file_path)
            
            data = [
                ["Header1", "Header2", "Header3"],
                ["Row1-1", "Row1-2", "Row1-3"],
                ["Row2-1", "Row2-2", "Row2-3"]
            ]
            
            result = create_table_with_data(
                file_path=file_path,
                rows=3,
                cols=3,
                data=data,
                header_bold=True
            )
            result_data = json.loads(result)
            
            assert result_data.get("success") == True

    def test_read_tables(self):
        """Test reading tables from document."""
        with tempfile.TemporaryDirectory() as tmpdir:
            file_path = os.path.join(tmpdir, "test.docx")
            create_new_document(file_path)
            
            data = [["A", "B"], ["1", "2"]]
            create_table_with_data(file_path, 2, 2, data)
            
            result = read_tables(file_path)
            tables_data = json.loads(result)
            
            assert "tables" in tables_data
            assert tables_data.get("total_tables", 0) >= 1

    def test_update_table_cell(self):
        """Test updating a table cell."""
        with tempfile.TemporaryDirectory() as tmpdir:
            file_path = os.path.join(tmpdir, "test.docx")
            create_new_document(file_path)
            
            data = [["A", "B"], ["1", "2"]]
            create_table_with_data(file_path, 2, 2, data)
            
            result = update_table_cell(
                file_path=file_path,
                table_index=0,
                row_index=1,
                col_index=0,
                new_text="Updated"
            )
            result_data = json.loads(result)
            
            assert result_data.get("success") == True

    def test_add_table_row(self):
        """Test adding a row to a table."""
        with tempfile.TemporaryDirectory() as tmpdir:
            file_path = os.path.join(tmpdir, "test.docx")
            create_new_document(file_path)
            
            data = [["A", "B"], ["1", "2"]]
            create_table_with_data(file_path, 2, 2, data)
            
            result = add_table_row(
                file_path=file_path,
                table_index=0,
                row_data=["3", "4"]
            )
            result_data = json.loads(result)
            
            assert result_data.get("success") == True

    def test_delete_table_row(self):
        """Test deleting a row from a table."""
        with tempfile.TemporaryDirectory() as tmpdir:
            file_path = os.path.join(tmpdir, "test.docx")
            create_new_document(file_path)
            
            data = [["A", "B"], ["1", "2"], ["3", "4"]]
            create_table_with_data(file_path, 3, 2, data)
            
            result = delete_table_row(
                file_path=file_path,
                table_index=0,
                row_index=1
            )
            result_data = json.loads(result)
            
            assert result_data.get("success") == True


class TestImageOperations:
    """Tests for image-related operations."""

    def test_get_images_info_empty_doc(self):
        """Test getting image info from document without images."""
        with tempfile.TemporaryDirectory() as tmpdir:
            file_path = os.path.join(tmpdir, "test.docx")
            create_new_document(file_path)
            
            result = get_images_info(file_path)
            data = json.loads(result)
            
            assert "total_images" in data
            assert data.get("total_images") == 0


class TestEdgeCases:
    """Tests for edge cases and error handling."""

    def test_replace_paragraph_invalid_index(self):
        """Test replacing paragraph with invalid index."""
        with tempfile.TemporaryDirectory() as tmpdir:
            file_path = os.path.join(tmpdir, "test.docx")
            create_new_document(file_path)
            add_formatted_paragraph(file_path, "Only one paragraph")
            
            result = replace_paragraph(
                file_path=file_path,
                paragraph_index=999,
                new_text="New text"
            )
            result_data = json.loads(result)
            
            # Should return error or handle gracefully
            assert "error" in result_data or "Error" in str(result_data)

    def test_delete_table_invalid_index(self):
        """Test deleting table with invalid index."""
        with tempfile.TemporaryDirectory() as tmpdir:
            file_path = os.path.join(tmpdir, "test.docx")
            create_new_document(file_path)
            
            result = delete_table(
                file_path=file_path,
                table_index=0
            )
            result_data = json.loads(result)
            
            assert "error" in result_data


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
