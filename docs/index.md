# LlamaDocx

A powerful Python toolkit for working with Microsoft Word documents.

## Overview

LlamaDocx provides a high-level interface for creating, modifying, and processing Word documents with features like template processing, text search and replace, format conversion, and more.

## Features

- **Document Processing**: Create, read, and modify Word documents with an intuitive API
- **Template System**: Fill Word document templates with data from JSON or Python dictionaries
- **Text Operations**: Search, replace, and find similar text using regex or fuzzy matching
- **Format Conversion**: Convert between Word and other formats (markdown, HTML, etc.) using pandoc
- **Image Handling**: Process and manipulate images within Word documents
- **Command Line Interface**: Access key features through a convenient CLI

## Installation

```bash
pip install llamadocx
```

For development installation with extra tools:

```bash
pip install llamadocx[dev]
```

For conversion capabilities:

```bash
pip install llamadocx[convert]
```

For all features:

```bash
pip install llamadocx[full]
```

## Quick Example

```python
from llamadocx import Document, Template

# Create a new document
doc = Document()
doc.add_heading("Hello World", level=1)
doc.add_paragraph("This is a test document.")
doc.save("test.docx")

# Process a template
template = Template("template.docx")
data = {
    "title": "My Report",
    "author": "John Doe",
    "sections": [
        {"heading": "Introduction", "content": "..."},
        {"heading": "Methods", "content": "..."},
    ]
}
template.merge_fields(data)
template.save("report.docx")

# Search and replace text
doc = Document("document.docx")
matches = doc.search("pattern", use_regex=True)
doc.replace("old text", "new text")
doc.save()
```

## Command Line Example

```bash
# Convert documents
llamadocx convert input.md -o output.docx
llamadocx convert document.docx --to-format markdown -o output.md

# Search for text
llamadocx search document.docx "pattern" --regex
llamadocx search document.docx "query" --similar

# Replace text
llamadocx replace document.docx "old" "new" --ignore-case

# Process templates
llamadocx template template.docx data.json -o output.docx
``` # Documentation updates
