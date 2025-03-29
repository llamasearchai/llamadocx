"""
LlamaDocx - A powerful toolkit for working with Microsoft Word documents.

LlamaDocx provides tools for document creation, editing, conversion, templating,
and more, with both a Python API and a command-line interface.
"""

__version__ = "0.1.0"

# Import the Document class from python-docx
# This allows users to import Document directly from llamadocx
from docx import Document

# Import core functionality
from llamadocx.template import Template
from llamadocx.image import Image, ImageProcessor
from llamadocx.metadata import (
    get_metadata, set_metadata, set_title, set_author, 
    set_subject, set_keywords, set_category, set_comments
)
from llamadocx.search import search_text, replace_text, find_similar_text
from llamadocx.convert import (
    convert_to_docx, convert_from_docx, docx_to_pdf, docx_to_markdown,
    docx_to_html, md_to_docx, html_to_docx
)

# Define what's available when importing * from llamadocx
__all__ = [
    # Classes
    'Document',
    'Template',
    'Image',
    'ImageProcessor',
    
    # Metadata functions
    'get_metadata',
    'set_metadata',
    'set_title',
    'set_author',
    'set_subject',
    'set_keywords',
    'set_category',
    'set_comments',
    
    # Search functions
    'search_text',
    'replace_text',
    'find_similar_text',
    
    # Convert functions
    'convert_to_docx',
    'convert_from_docx', 
    'docx_to_pdf', 
    'docx_to_markdown',
    'docx_to_html', 
    'md_to_docx', 
    'html_to_docx',
] 