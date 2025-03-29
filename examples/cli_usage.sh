#!/bin/bash
# CLI Usage Examples for LlamaDocx
# This script demonstrates various ways to use the LlamaDocx command-line interface

# First, make sure LlamaDocx is installed
# pip install llamadocx

# Set up example directory
EXAMPLE_DIR="llamadocx_cli_examples"
mkdir -p "$EXAMPLE_DIR"
cd "$EXAMPLE_DIR"

echo "Creating example files in $EXAMPLE_DIR"
echo

# 1. Show version information
echo "=== LlamaDocx Version Information ==="
llamadocx --version
echo

# 2. Create a new document
echo "=== Creating a new document ==="
llamadocx create sample.docx --title "Sample Document" --author "LlamaDocx User"
echo "Created sample.docx"
echo

# 3. Create a minimal Markdown file for conversion
echo "=== Creating sample Markdown file ==="
cat > sample.md << 'EOL'
# Sample Markdown Document

## Introduction

This is a sample document created with **Markdown**.

## Features

- Easy to read and write
- Plain text format
- Converts to many formats

## Conclusion

LlamaDocx makes it easy to work with Word documents.
EOL
echo "Created sample.md"
echo

# 4. Convert Markdown to DOCX
echo "=== Converting Markdown to DOCX ==="
llamadocx convert sample.md sample_from_md.docx
echo "Converted sample.md to sample_from_md.docx"
echo

# 5. Extract text from DOCX
echo "=== Extracting text from DOCX ==="
llamadocx extract sample_from_md.docx --format text --output extracted_text.txt
echo "Extracted text to extracted_text.txt"
echo "Content of extracted_text.txt:"
cat extracted_text.txt
echo

# 6. Extract as JSON
echo "=== Extracting as JSON ==="
llamadocx extract sample_from_md.docx --format json --output extracted.json
echo "Extracted to extracted.json"
echo

# 7. Search for text in document
echo "=== Searching document ==="
llamadocx search sample_from_md.docx "Markdown" --context 1
echo

# 8. Replace text in document
echo "=== Replacing text ==="
llamadocx replace sample_from_md.docx "Markdown" "Markdown format" --count --output modified.docx
echo "Created modified.docx with replacements"
echo

# 9. Get document metadata
echo "=== Document metadata ==="
llamadocx meta --get sample.docx
echo

# 10. Create a template document
echo "=== Creating template document ==="
cat > template.docx.json << 'EOL'
{
    "title": "Quarterly Report",
    "author": "Finance Team",
    "date": "2024-03-31",
    "department": "Finance",
    "content": "This report contains financial information for Q1 2024."
}
EOL

# Create template document
llamadocx create template.docx --title "Report Template"
echo "Created template.docx"

# Process template
echo "=== Processing template ==="
llamadocx template template.docx template.docx.json processed_template.docx
echo "Processed template to processed_template.docx"
echo

# 11. Convert DOCX to Markdown
echo "=== Converting DOCX to Markdown ==="
llamadocx convert sample.docx sample_to_md.md 
echo "Converted sample.docx to sample_to_md.md"
echo

# 12. List all available commands
echo "=== LlamaDocx Help ==="
llamadocx --help
echo

# Cleanup
read -p "Do you want to clean up example files? (y/n) " -n 1 -r
echo
if [[ $REPLY =~ ^[Yy]$ ]]
then
    cd ..
    rm -rf "$EXAMPLE_DIR"
    echo "Removed $EXAMPLE_DIR"
fi 