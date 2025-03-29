"""
Functions for working with document metadata.

This module provides functions for reading, writing, and updating document
metadata properties such as title, author, keywords, and other core properties.
"""

from pathlib import Path
from datetime import datetime
from typing import Optional, Union, Dict, Any, List

from .document import Document


def get_metadata(doc: Document) -> Dict[str, Any]:
    """
    Get all metadata from a document.

    Args:
        doc (Document): The document to extract metadata from

    Returns:
        Dict[str, Any]: Dictionary of metadata properties
    """
    return doc.core_properties


def set_metadata(
    doc: Document,
    properties: Dict[str, Any]
) -> None:
    """
    Set metadata properties for a document.

    Args:
        doc (Document): The document to update
        properties (Dict[str, Any]): Dictionary of metadata properties to set
    """
    core_props = doc.doc.core_properties
    
    # Map common properties to their respective attributes
    for key, value in properties.items():
        if hasattr(core_props, key):
            setattr(core_props, key, value)


def set_title(doc: Document, title: str) -> None:
    """
    Set the document title.

    Args:
        doc (Document): The document to update
        title (str): Title to set
    """
    doc.doc.core_properties.title = title


def set_author(doc: Document, author: str) -> None:
    """
    Set the document author.

    Args:
        doc (Document): The document to update
        author (str): Author to set
    """
    doc.doc.core_properties.author = author


def set_subject(doc: Document, subject: str) -> None:
    """
    Set the document subject.

    Args:
        doc (Document): The document to update
        subject (str): Subject to set
    """
    doc.doc.core_properties.subject = subject


def set_keywords(doc: Document, keywords: Union[str, List[str]]) -> None:
    """
    Set the document keywords.

    Args:
        doc (Document): The document to update
        keywords (Union[str, List[str]]): Keywords to set.
            If a list is provided, it will be joined with semicolons.
    """
    if isinstance(keywords, list):
        keywords = "; ".join(keywords)
    
    doc.doc.core_properties.keywords = keywords


def set_category(doc: Document, category: str) -> None:
    """
    Set the document category.

    Args:
        doc (Document): The document to update
        category (str): Category to set
    """
    doc.doc.core_properties.category = category


def set_comments(doc: Document, comments: str) -> None:
    """
    Set the document comments.

    Args:
        doc (Document): The document to update
        comments (str): Comments to set
    """
    doc.doc.core_properties.comments = comments


def set_created_time(
    doc: Document,
    created: Optional[Union[datetime, str]] = None
) -> None:
    """
    Set the document creation time.

    Args:
        doc (Document): The document to update
        created (Union[datetime, str, None]): Creation time to set.
            If None, current time is used.
    """
    if created is None:
        created = datetime.now()
    elif isinstance(created, str):
        created = datetime.fromisoformat(created)
    
    doc.doc.core_properties.created = created


def set_last_modified_time(
    doc: Document,
    modified: Optional[Union[datetime, str]] = None
) -> None:
    """
    Set the document last modified time.

    Args:
        doc (Document): The document to update
        modified (Union[datetime, str, None]): Modification time to set.
            If None, current time is used.
    """
    if modified is None:
        modified = datetime.now()
    elif isinstance(modified, str):
        modified = datetime.fromisoformat(modified)
    
    doc.doc.core_properties.modified = modified


def update_metadata_from_file(
    doc: Document,
    path: Union[str, Path]
) -> None:
    """
    Update document metadata from a file.
    
    The file should be in JSON format containing key-value pairs for metadata.

    Args:
        doc (Document): The document to update
        path (Union[str, Path]): Path to the metadata file
    
    Raises:
        FileNotFoundError: If the file doesn't exist
        ValueError: If the file isn't valid JSON
    """
    import json
    from pathlib import Path
    
    path = Path(path) if isinstance(path, str) else path
    
    if not path.exists():
        raise FileNotFoundError(f"Metadata file not found: {path}")
    
    try:
        with open(path, 'r', encoding='utf-8') as f:
            metadata = json.load(f)
    except json.JSONDecodeError as e:
        raise ValueError(f"Invalid JSON in metadata file: {e}")
    
    set_metadata(doc, metadata)


def extract_metadata_to_file(
    doc: Document,
    path: Union[str, Path]
) -> None:
    """
    Extract document metadata to a file in JSON format.

    Args:
        doc (Document): The document to extract metadata from
        path (Union[str, Path]): Path to save the metadata file
    """
    import json
    from pathlib import Path
    
    path = Path(path) if isinstance(path, str) else path
    
    # Get metadata
    metadata = get_metadata(doc)
    
    # Convert datetime objects to ISO format strings
    for key, value in metadata.items():
        if isinstance(value, datetime):
            metadata[key] = value.isoformat()
    
    # Write to file
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(metadata, f, indent=2) 