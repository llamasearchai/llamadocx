"""
Functions for searching and replacing text in Word documents.

This module provides functions for searching and replacing text in Word
documents, with support for regular expressions and case sensitivity.
"""

import re
from typing import List, Dict, Any, Optional, Union, Tuple

from .document import Document
from .paragraph import Paragraph
from .table import Table, Cell


def search_text(
    doc: Document,
    pattern: str,
    regex: bool = False,
    case_sensitive: bool = False
) -> List[Dict[str, Any]]:
    """
    Search for text in a document.

    Args:
        doc (Document): The document to search
        pattern (str): Text pattern to search for
        regex (bool): Whether to treat pattern as regex
        case_sensitive (bool): Whether to match case

    Returns:
        List[Dict[str, Any]]: List of matches with location info
    """
    matches = []

    # Compile pattern
    if regex:
        flags = 0 if case_sensitive else re.IGNORECASE
        pattern_obj = re.compile(pattern, flags)
    else:
        pattern = re.escape(pattern)
        flags = 0 if case_sensitive else re.IGNORECASE
        pattern_obj = re.compile(pattern, flags)

    # Search paragraphs
    for para_idx, paragraph in enumerate(doc.paragraphs):
        para_matches = _search_paragraph(
            paragraph,
            pattern_obj,
            location={"type": "paragraph", "index": para_idx}
        )
        matches.extend(para_matches)

    # Search tables
    for table_idx, table in enumerate(doc.tables):
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                for para_idx, paragraph in enumerate(cell.paragraphs):
                    para_matches = _search_paragraph(
                        paragraph,
                        pattern_obj,
                        location={
                            "type": "table",
                            "table_index": table_idx,
                            "row_index": row_idx,
                            "column_index": col_idx,
                            "paragraph_index": para_idx
                        }
                    )
                    matches.extend(para_matches)

    return matches


def _search_paragraph(
    paragraph: Paragraph,
    pattern: re.Pattern,
    location: Dict[str, Any]
) -> List[Dict[str, Any]]:
    """
    Search for matches in a paragraph.

    Args:
        paragraph (Paragraph): The paragraph to search
        pattern (re.Pattern): Compiled regex pattern
        location (Dict[str, Any]): Location info for the paragraph

    Returns:
        List[Dict[str, Any]]: List of matches with location info
    """
    matches = []
    text = paragraph.text

    for match in pattern.finditer(text):
        matches.append({
            "text": match.group(0),
            "start": match.start(),
            "end": match.end(),
            "location": location,
            "context": _get_context(text, match.start(), match.end())
        })

    return matches


def _get_context(
    text: str,
    start: int,
    end: int,
    context_chars: int = 50
) -> Dict[str, str]:
    """
    Get context around a match.

    Args:
        text (str): The full text
        start (int): Match start position
        end (int): Match end position
        context_chars (int): Number of context characters

    Returns:
        Dict[str, str]: Before and after context
    """
    before_start = max(0, start - context_chars)
    after_end = min(len(text), end + context_chars)

    return {
        "before": text[before_start:start],
        "after": text[end:after_end]
    }


def replace_text(
    doc: Document,
    pattern: str,
    replacement: str,
    regex: bool = False,
    case_sensitive: bool = False
) -> int:
    """
    Replace text in a document.

    Args:
        doc (Document): The document to modify
        pattern (str): Text pattern to search for
        replacement (str): Replacement text
        regex (bool): Whether to treat pattern as regex
        case_sensitive (bool): Whether to match case

    Returns:
        int: Number of replacements made
    """
    # Compile pattern
    if regex:
        flags = 0 if case_sensitive else re.IGNORECASE
        pattern_obj = re.compile(pattern, flags)
    else:
        pattern = re.escape(pattern)
        flags = 0 if case_sensitive else re.IGNORECASE
        pattern_obj = re.compile(pattern, flags)

    replacements = 0

    # Replace in paragraphs
    for paragraph in doc.paragraphs:
        count = _replace_paragraph(paragraph, pattern_obj, replacement)
        replacements += count

    # Replace in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    count = _replace_paragraph(paragraph, pattern_obj, replacement)
                    replacements += count

    return replacements


def _replace_paragraph(
    paragraph: Paragraph,
    pattern: re.Pattern,
    replacement: str
) -> int:
    """
    Replace matches in a paragraph.

    Args:
        paragraph (Paragraph): The paragraph to modify
        pattern (re.Pattern): Compiled regex pattern
        replacement (str): Replacement text

    Returns:
        int: Number of replacements made
    """
    text = paragraph.text
    new_text, count = pattern.subn(replacement, text)

    if count > 0:
        # Clear existing runs
        paragraph.clear()

        # Add new text
        paragraph.add_run(new_text)

    return count


def find_similar_text(
    doc: Document,
    query: str,
    threshold: float = 0.8,
    max_results: int = 10
) -> List[Dict[str, Any]]:
    """
    Find text similar to a query using fuzzy matching.

    Args:
        doc (Document): The document to search
        query (str): Text to search for
        threshold (float): Similarity threshold (0-1)
        max_results (int): Maximum number of results

    Returns:
        List[Dict[str, Any]]: List of matches with similarity scores
    """
    try:
        from rapidfuzz import fuzz, process
    except ImportError:
        raise ImportError(
            "rapidfuzz is required for fuzzy text matching. "
            "Install it with: pip install rapidfuzz"
        )

    matches = []

    # Search paragraphs
    for para_idx, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text
        if not text.strip():
            continue

        score = fuzz.ratio(query.lower(), text.lower()) / 100
        if score >= threshold:
            matches.append({
                "text": text,
                "score": score,
                "location": {
                    "type": "paragraph",
                    "index": para_idx
                }
            })

    # Search tables
    for table_idx, table in enumerate(doc.tables):
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                for para_idx, paragraph in enumerate(cell.paragraphs):
                    text = paragraph.text
                    if not text.strip():
                        continue

                    score = fuzz.ratio(query.lower(), text.lower()) / 100
                    if score >= threshold:
                        matches.append({
                            "text": text,
                            "score": score,
                            "location": {
                                "type": "table",
                                "table_index": table_idx,
                                "row_index": row_idx,
                                "column_index": col_idx,
                                "paragraph_index": para_idx
                            }
                        })

    # Sort by score and limit results
    matches.sort(key=lambda x: x["score"], reverse=True)
    return matches[:max_results] 