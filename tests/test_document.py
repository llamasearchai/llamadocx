"""
Tests for the Document class.
"""

import os
import pytest
from pathlib import Path
from llamadocx import Document


@pytest.fixture
def temp_doc(tmp_path):
    """Create a temporary document for testing."""
    doc = Document()
    doc.add_heading("Test Document", level=1)
    doc.add_paragraph("This is a test paragraph.")
    path = tmp_path / "test.docx"
    doc.save(path)
    return path


def test_create_document():
    """Test creating a new document."""
    doc = Document()
    assert doc is not None
    assert hasattr(doc, '_document')


def test_add_heading(temp_doc):
    """Test adding a heading to a document."""
    doc = Document(temp_doc)
    doc.add_heading("New Section", level=2)
    text = doc.get_text()
    assert "New Section" in text


def test_add_paragraph(temp_doc):
    """Test adding a paragraph to a document."""
    doc = Document(temp_doc)
    test_text = "This is a new paragraph for testing."
    doc.add_paragraph(test_text)
    text = doc.get_text()
    assert test_text in text


def test_add_table(temp_doc):
    """Test adding a table to a document."""
    doc = Document(temp_doc)
    table = doc.add_table(rows=2, cols=2)
    
    # Fill table with test data
    data = [["A1", "B1"], ["A2", "B2"]]
    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            cell.text = data[i][j]
    
    # Verify table contents
    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            assert cell.text == data[i][j]


def test_save_document(tmp_path):
    """Test saving a document."""
    doc = Document()
    doc.add_heading("Test Save", level=1)
    
    # Save document
    path = tmp_path / "save_test.docx"
    doc.save(path)
    
    # Verify file exists and can be opened
    assert path.exists()
    doc2 = Document(path)
    text = doc2.get_text()
    assert "Test Save" in text


def test_get_text(temp_doc):
    """Test extracting text from a document."""
    doc = Document(temp_doc)
    text = doc.get_text()
    assert "Test Document" in text
    assert "This is a test paragraph." in text


def test_document_properties(temp_doc):
    """Test document properties."""
    doc = Document(temp_doc)
    
    # Test basic properties
    assert doc.paragraphs is not None
    assert len(doc.paragraphs) > 0
    assert doc.tables is not None
    
    # Add and test section properties
    doc.add_section()
    assert len(doc.sections) > 0


def test_invalid_file():
    """Test handling of invalid file paths."""
    with pytest.raises(FileNotFoundError):
        Document("nonexistent.docx")


def test_style_operations(temp_doc):
    """Test operations with document styles."""
    doc = Document(temp_doc)
    
    # Add paragraph with style
    para = doc.add_paragraph("Styled text", style="Heading 1")
    assert para.style.name == "Heading 1"
    
    # Change style
    para.style = "Normal"
    assert para.style.name == "Normal" 