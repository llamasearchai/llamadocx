# LlamaDocx

A powerful Python toolkit for working with Microsoft Word documents. LlamaDocx provides a high-level interface for creating, modifying, and processing Word documents with features like template processing, text search and replace, format conversion, and more.

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

### Requirements

- Python 3.8 or higher
- pandoc (for document conversion)
- Microsoft Word or compatible software (for viewing .docx files)

## Quick Start

### Python API

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

### Command Line Interface

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
```

## Documentation

For detailed documentation, visit [llamadocx.readthedocs.io](https://llamasearch.ai

### Examples

Check the `examples` directory for more usage examples:

- Basic document creation and modification
- Template processing with different data structures
- Advanced search and replace operations
- Format conversion examples
- Image processing in documents

## Development

To set up the development environment:

```bash
git clone https://llamasearch.ai
cd llamadocx
pip install -e .[dev]
```

Run tests:

```bash
pytest
```

## Contributing

Contributions are welcome! Please check our [Contributing Guidelines](CONTRIBUTING.md) for details.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Credits

LlamaDocx is developed and maintained by [LlamaSearch.AI](https://llamasearch.ai 
# Updated in commit 1 - 2025-04-04 17:21:33

# Updated in commit 9 - 2025-04-04 17:21:33

# Updated in commit 17 - 2025-04-04 17:21:34

# Updated in commit 25 - 2025-04-04 17:21:35

# Updated in commit 1 - 2025-04-05 14:30:42

# Updated in commit 9 - 2025-04-05 14:30:42

# Updated in commit 17 - 2025-04-05 14:30:42

# Updated in commit 25 - 2025-04-05 14:30:42

# Updated in commit 1 - 2025-04-05 15:17:09

# Updated in commit 9 - 2025-04-05 15:17:09

# Updated in commit 17 - 2025-04-05 15:17:09

# Updated in commit 25 - 2025-04-05 15:17:10

# Updated in commit 1 - 2025-04-05 15:47:55

# Updated in commit 9 - 2025-04-05 15:47:55
