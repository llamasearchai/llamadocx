#!/usr/bin/env python3
"""
Format conversion example for llamadocx.

This example demonstrates the format conversion capabilities of llamadocx,
including converting between different document formats (DOCX, PDF, Markdown, HTML).
"""

import os
import sys
from pathlib import Path
from llamadocx import Document
from llamadocx.utils import ensure_directory


def create_sample_markdown():
    """Create a sample Markdown file for conversion."""
    md_content = """# Sample Markdown Document

## Introduction

This is a sample Markdown document created for demonstrating format conversion capabilities of LlamaDocx.

## Features

### Rich Text Support

Markdown supports various text formatting options:
- **Bold text** for emphasis
- *Italic text* for subtle emphasis
- ~~Strikethrough~~ for marking deleted content
- `inline code` for technical references

### Lists

Ordered list:
1. First item
2. Second item
3. Third item with sub-items:
   a. Sub-item 1
   b. Sub-item 2

Unordered list:
- Main point
- Another point
  - Sub-point
  - Another sub-point
- Final point

### Tables

| Name | Role | Department |
|------|------|------------|
| John Doe | Developer | Engineering |
| Jane Smith | Designer | UX/UI |
| Bob Johnson | Manager | Product |

### Code Blocks

```python
def hello_world():
    print("Hello, World!")
    return True
```

## Images

![Placeholder for image](path/to/image.jpg)

## Links

Visit [LlamaSearch.AI](https://llamasearch.ai) for more information.

## Conclusion

This document demonstrates the various elements that can be converted between different formats.
"""

    # Save the markdown file
    md_path = Path("sample_document.md")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(md_content)
    
    print(f"Sample Markdown file created: {md_path.absolute()}")
    return md_path


def create_sample_html():
    """Create a sample HTML file for conversion."""
    html_content = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sample HTML Document</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            line-height: 1.6;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
        }
        h1 {
            color: #2c3e50;
            border-bottom: 2px solid #ecf0f1;
            padding-bottom: 10px;
        }
        h2 {
            color: #3498db;
            margin-top: 30px;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin: 20px 0;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }
        th {
            background-color: #f2f2f2;
        }
        .highlight {
            background-color: #ffffcc;
            padding: 2px 5px;
            border-radius: 3px;
        }
        code {
            background-color: #f8f8f8;
            padding: 2px 5px;
            border: 1px solid #ddd;
            border-radius: 3px;
            font-family: "Courier New", monospace;
        }
        pre {
            background-color: #f8f8f8;
            padding: 15px;
            border: 1px solid #ddd;
            border-radius: 3px;
            overflow-x: auto;
        }
    </style>
</head>
<body>
    <h1>Sample HTML Document</h1>
    
    <h2>Introduction</h2>
    <p>This is a sample HTML document created for demonstrating format conversion capabilities of LlamaDocx.</p>
    
    <h2>Features</h2>
    
    <h3>Rich Text Support</h3>
    <p>HTML supports various text formatting options:</p>
    <ul>
        <li><strong>Bold text</strong> for emphasis</li>
        <li><em>Italic text</em> for subtle emphasis</li>
        <li><del>Strikethrough</del> for marking deleted content</li>
        <li><code>inline code</code> for technical references</li>
        <li><span class="highlight">Highlighted text</span> for important notes</li>
    </ul>
    
    <h3>Lists</h3>
    
    <p>Ordered list:</p>
    <ol>
        <li>First item</li>
        <li>Second item</li>
        <li>Third item with sub-items:
            <ol type="a">
                <li>Sub-item 1</li>
                <li>Sub-item 2</li>
            </ol>
        </li>
    </ol>
    
    <p>Unordered list:</p>
    <ul>
        <li>Main point</li>
        <li>Another point
            <ul>
                <li>Sub-point</li>
                <li>Another sub-point</li>
            </ul>
        </li>
        <li>Final point</li>
    </ul>
    
    <h3>Tables</h3>
    
    <table>
        <thead>
            <tr>
                <th>Name</th>
                <th>Role</th>
                <th>Department</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>John Doe</td>
                <td>Developer</td>
                <td>Engineering</td>
            </tr>
            <tr>
                <td>Jane Smith</td>
                <td>Designer</td>
                <td>UX/UI</td>
            </tr>
            <tr>
                <td>Bob Johnson</td>
                <td>Manager</td>
                <td>Product</td>
            </tr>
        </tbody>
    </table>
    
    <h3>Code Blocks</h3>
    
    <pre><code>def hello_world():
    print("Hello, World!")
    return True</code></pre>
    
    <h2>Images</h2>
    
    <p><img src="path/to/image.jpg" alt="Placeholder for image"></p>
    
    <h2>Links</h2>
    
    <p>Visit <a href="https://llamasearch.ai">LlamaSearch.AI</a> for more information.</p>
    
    <h2>Conclusion</h2>
    
    <p>This document demonstrates the various elements that can be converted between different formats.</p>
</body>
</html>"""

    # Save the HTML file
    html_path = Path("sample_document.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    
    print(f"Sample HTML file created: {html_path.absolute()}")
    return html_path


def convert_md_to_docx(md_path):
    """Convert a Markdown file to DOCX format."""
    print(f"\nConverting Markdown to DOCX: {md_path}")
    
    try:
        from llamadocx.convert import md_to_docx
        
        # Output path
        output_path = md_path.with_suffix('.docx')
        
        # Convert
        md_to_docx(md_path, output_path)
        
        print(f"Conversion successful. DOCX file saved to: {output_path.absolute()}")
        return output_path
    except ImportError as e:
        print(f"Error: {e}")
        print("Markdown conversion requires the 'markdown' and 'python-docx' packages.")
        print("Install them with: pip install markdown python-docx")
        return None


def convert_html_to_docx(html_path):
    """Convert an HTML file to DOCX format."""
    print(f"\nConverting HTML to DOCX: {html_path}")
    
    try:
        from llamadocx.convert import html_to_docx
        
        # Output path
        output_path = html_path.with_suffix('.docx')
        
        # Convert
        html_to_docx(html_path, output_path)
        
        print(f"Conversion successful. DOCX file saved to: {output_path.absolute()}")
        return output_path
    except ImportError as e:
        print(f"Error: {e}")
        print("HTML conversion requires the 'python-docx' and 'beautifulsoup4' packages.")
        print("Install them with: pip install python-docx beautifulsoup4")
        return None


def convert_docx_to_pdf(docx_path):
    """Convert a DOCX file to PDF format."""
    print(f"\nConverting DOCX to PDF: {docx_path}")
    
    try:
        from llamadocx.convert import docx_to_pdf
        
        # Output path
        output_path = docx_path.with_suffix('.pdf')
        
        # Convert
        docx_to_pdf(docx_path, output_path)
        
        print(f"Conversion successful. PDF file saved to: {output_path.absolute()}")
        return output_path
    except ImportError as e:
        print(f"Error: {e}")
        print("PDF conversion requires additional dependencies.")
        print("Install them with: pip install docx2pdf")
        return None


def convert_docx_to_markdown(docx_path):
    """Convert a DOCX file to Markdown format."""
    print(f"\nConverting DOCX to Markdown: {docx_path}")
    
    try:
        from llamadocx.convert import docx_to_markdown
        
        # Output path
        output_path = docx_path.parent / f"{docx_path.stem}_converted.md"
        
        # Convert
        docx_to_markdown(docx_path, output_path)
        
        print(f"Conversion successful. Markdown file saved to: {output_path.absolute()}")
        return output_path
    except ImportError as e:
        print(f"Error: {e}")
        print("Markdown conversion requires additional dependencies.")
        print("Install them with: pip install pypandoc")
        return None


def convert_docx_to_html(docx_path):
    """Convert a DOCX file to HTML format."""
    print(f"\nConverting DOCX to HTML: {docx_path}")
    
    try:
        from llamadocx.convert import docx_to_html
        
        # Output path
        output_path = docx_path.parent / f"{docx_path.stem}_converted.html"
        
        # Convert
        docx_to_html(docx_path, output_path)
        
        print(f"Conversion successful. HTML file saved to: {output_path.absolute()}")
        return output_path
    except ImportError as e:
        print(f"Error: {e}")
        print("HTML conversion requires additional dependencies.")
        print("Install them with: pip install pypandoc")
        return None


def create_output_directory():
    """Create a directory for output files."""
    output_dir = Path("conversion_output")
    ensure_directory(output_dir)
    return output_dir


def main():
    """Main function for format conversion example."""
    # Create output directory
    output_dir = create_output_directory()
    
    # Change working directory to output directory
    os.chdir(output_dir)
    
    # Create sample files
    md_path = create_sample_markdown()
    html_path = create_sample_html()
    
    # Perform conversions
    md_docx_path = convert_md_to_docx(md_path)
    html_docx_path = convert_html_to_docx(html_path)
    
    # If conversions succeeded, convert to other formats
    if md_docx_path:
        convert_docx_to_pdf(md_docx_path)
        convert_docx_to_markdown(md_docx_path)
        convert_docx_to_html(md_docx_path)
    
    if html_docx_path:
        convert_docx_to_pdf(html_docx_path)
    
    print("\nFormat conversion example completed. Check the 'conversion_output' directory for files.")


if __name__ == "__main__":
    main() 