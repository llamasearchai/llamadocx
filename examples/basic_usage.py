#!/usr/bin/env python3
"""
Basic usage examples for llamadocx.

This example demonstrates the core functionality of the llamadocx library,
including document creation, modification, and text operations.
"""

import os
from pathlib import Path
from llamadocx import Document, Template, convert_to_docx
from llamadocx.metadata import set_author, set_title


def create_document_example():
    """Create a simple document with various elements."""
    print("Creating a new document...")
    
    # Create a new document
    doc = Document()
    
    # Add a heading
    doc.add_heading("LlamaDocx Example Document", level=1)
    
    # Add some paragraphs
    doc.add_paragraph("This is a sample document created with LlamaDocx.")
    
    p = doc.add_paragraph("You can format text in paragraphs. ")
    p.add_run("This text is bold.").bold = True
    p.add_run(" This text is italic.").italic = True
    p.add_run(" This text is underlined.").underline = True
    
    # Add a heading for a new section
    doc.add_heading("Tables", level=2)
    
    # Add a paragraph
    doc.add_paragraph("LlamaDocx makes it easy to create tables:")
    
    # Add a table
    table = doc.add_table(rows=3, cols=3)
    
    # Add header row
    header_cells = table.rows[0].cells
    header_cells[0].text = "Header 1"
    header_cells[1].text = "Header 2"
    header_cells[2].text = "Header 3"
    
    # Add data rows
    table.rows[1].cells[0].text = "Row 1, Cell 1"
    table.rows[1].cells[1].text = "Row 1, Cell 2"
    table.rows[1].cells[2].text = "Row 1, Cell 3"
    
    table.rows[2].cells[0].text = "Row 2, Cell 1"
    table.rows[2].cells[1].text = "Row 2, Cell 2"
    table.rows[2].cells[2].text = "Row 2, Cell 3"
    
    # Add metadata
    set_title(doc, "LlamaDocx Example")
    set_author(doc, "LlamaSearch.AI")
    
    # Save the document
    output_path = Path("basic_example.docx")
    doc.save(output_path)
    print(f"Document saved to: {output_path.absolute()}")
    
    return output_path


def modify_document_example(doc_path):
    """Modify an existing document."""
    print(f"\nModifying document: {doc_path}")
    
    # Open the document
    doc = Document(doc_path)
    
    # Add a new section
    doc.add_heading("Images", level=2)
    
    # Add a paragraph
    doc.add_paragraph("You can also add images to your documents.")
    
    # Add an image (if an image file exists)
    image_path = Path("example_image.png")
    if not image_path.exists():
        # Create a simple text file as a placeholder
        with open(image_path, "w") as f:
            f.write("This is a placeholder for an image file.")
        print(f"Note: Created placeholder file {image_path} since no actual image exists")
    
    try:
        doc.add_image(image_path, width=4)  # 4 inches wide
    except Exception as e:
        print(f"Note: Could not add image: {e}")
        doc.add_paragraph(f"[Image would be added here: {image_path}]")
    
    # Save the modified document
    output_path = Path("modified_example.docx")
    doc.save(output_path)
    print(f"Modified document saved to: {output_path.absolute()}")
    
    return output_path


def search_replace_example(doc_path):
    """Demonstrate search and replace functionality."""
    print(f"\nPerforming search and replace on: {doc_path}")
    
    # Open the document
    doc = Document(doc_path)
    
    # Search for text
    results = doc.search_text("LlamaDocx")
    print(f"Found {len(results)} occurrences of 'LlamaDocx'")
    
    # Replace text
    replacements = doc.replace_text("LlamaDocx", "LlamaDocx™")
    print(f"Made {replacements} replacements")
    
    # Save the document
    output_path = Path("search_replace_example.docx")
    doc.save(output_path)
    print(f"Document with replacements saved to: {output_path.absolute()}")
    
    return output_path


def template_example():
    """Demonstrate template functionality."""
    print("\nCreating and processing a template document...")
    
    # Create a template document
    doc = Document()
    doc.add_heading("{{title}}", level=1)
    doc.add_paragraph("Author: {{author}}")
    doc.add_paragraph("Date: {{date}}")
    
    # Add a repeating section
    doc.add_heading("Table of Contents", level=2)
    doc.add_paragraph("{{#sections}}")
    doc.add_paragraph("{{section_number}}. {{section_title}}")
    doc.add_paragraph("{{/sections}}")
    
    # Add content section
    doc.add_heading("Content", level=2)
    doc.add_paragraph("{{content}}")
    
    # Save the template
    template_path = Path("template_example.docx")
    doc.save(template_path)
    
    # Create a Template object from the document
    template = Template(template_path)
    
    # Prepare data for the template
    data = {
        "title": "Quarterly Report",
        "author": "John Smith",
        "date": "2024-03-31",
        "sections": [
            {"section_number": 1, "section_title": "Introduction"},
            {"section_number": 2, "section_title": "Financial Overview"},
            {"section_number": 3, "section_title": "Key Projects"},
            {"section_number": 4, "section_title": "Future Plans"},
        ],
        "content": "This is the main content of the document. It would include detailed information about each section listed in the table of contents."
    }
    
    # Merge data into the template
    template.merge_fields(data)
    
    # Save the processed template
    output_path = Path("filled_template.docx")
    template.save(output_path)
    print(f"Processed template saved to: {output_path.absolute()}")
    
    return output_path


def conversion_example():
    """Demonstrate format conversion functionality."""
    # Check if we have a document to convert
    md_content = """# Markdown Document
    
## Introduction

This is a sample markdown document created for the LlamaDocx example.

## Features

- Format conversion
- Document creation
- Template processing
- Text operations
"""
    
    md_path = Path("example.md")
    with open(md_path, "w") as f:
        f.write(md_content)
    
    print(f"\nConverting {md_path} to DOCX format...")
    
    # Convert markdown to docx
    try:
        output_path = convert_to_docx(md_path)
        print(f"Converted document saved to: {output_path.absolute()}")
    except ImportError:
        print("Note: Conversion requires pypandoc and pandoc. "
              "Skipping conversion example.")
        output_path = None
    
    return output_path


if __name__ == "__main__":
    # Create a basic document
    doc_path = create_document_example()
    
    # Modify the document
    modified_path = modify_document_example(doc_path)
    
    # Search and replace operations
    search_replace_example(modified_path)
    
    # Template operations
    template_example()
    
    # Format conversion
    conversion_example()
    
    print("\nAll examples completed. Check the current directory for output files.") # Example scripts
