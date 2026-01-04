"""Tests for base DOCX tool (markdown_to_word).

These tests verify that the markdown to Word conversion works correctly,
including headers, lists, tables, formatting, links, and block quotes.

Output files are saved to tests/output/docx/base/ directory for manual inspection.
"""

import os
import sys
from pathlib import Path
from io import BytesIO

# Add project root to path for imports
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest
from docx import Document

from docx_tools.base_docx_tool import markdown_to_word
from docx_tools.helpers import parse_inline_formatting

# Output directory for test files
OUTPUT_DIR = Path(__file__).parent / "output" / "docx" / "base"


@pytest.fixture(scope="module", autouse=True)
def setup_output_dir():
    """Create output directory if it doesn't exist."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    yield


def get_document_from_result(result: str) -> Document:
    """Load the generated document from the result path.

    Note: This only works with LOCAL upload strategy.
    """
    # Result is typically a path like "output/uuid.docx" or a URL
    if result.startswith("http"):
        pytest.skip("Cannot verify document content with remote upload strategy")

    # Extract path from result
    path = Path(result)
    if not path.exists():
        # Try relative to project root
        path = project_root / result

    if path.exists():
        return Document(path)
    else:
        pytest.skip(f"Cannot find output file: {result}")


def save_test_document(markdown: str, filename: str) -> str:
    """Helper to convert markdown and save for inspection."""
    result = markdown_to_word(markdown)

    # Copy the file to test output directory for inspection
    if not result.startswith("http"):
        src_path = Path(result)
        if not src_path.exists():
            src_path = project_root / result

        if src_path.exists():
            dest_path = OUTPUT_DIR / filename
            import shutil
            shutil.copy(src_path, dest_path)
            print(f"Saved: {dest_path}")

    return result


# =============================================================================
# Header Tests
# =============================================================================

class TestHeaders:
    """Tests for markdown headers conversion."""

    def test_h1_header(self):
        """Test H1 header conversion."""
        markdown = "# Main Title"
        result = save_test_document(markdown, "header_h1.docx")
        assert not result.startswith("Error")

    def test_h2_header(self):
        """Test H2 header conversion."""
        markdown = "## Section Title"
        result = save_test_document(markdown, "header_h2.docx")
        assert not result.startswith("Error")

    def test_h3_header(self):
        """Test H3 header conversion."""
        markdown = "### Subsection Title"
        result = save_test_document(markdown, "header_h3.docx")
        assert not result.startswith("Error")

    def test_multiple_headers(self):
        """Test document with multiple header levels."""
        markdown = """# Document Title

## Introduction

Some intro text here.

### Details

More details.

## Conclusion

Final thoughts.
"""
        result = save_test_document(markdown, "header_multiple.docx")
        assert not result.startswith("Error")

    def test_header_with_formatting(self):
        """Test header with inline formatting."""
        markdown = "# Title with **bold** and *italic*"
        result = save_test_document(markdown, "header_formatted.docx")
        assert not result.startswith("Error")


# =============================================================================
# List Tests
# =============================================================================

class TestLists:
    """Tests for markdown list conversion."""

    def test_unordered_list(self):
        """Test unordered (bullet) list."""
        markdown = """- First item
- Second item
- Third item
"""
        result = save_test_document(markdown, "list_unordered.docx")
        assert not result.startswith("Error")

    def test_ordered_list(self):
        """Test ordered (numbered) list."""
        markdown = """1. First item
2. Second item
3. Third item
"""
        result = save_test_document(markdown, "list_ordered.docx")
        assert not result.startswith("Error")

    def test_nested_list(self):
        """Test nested list items."""
        markdown = """- Main item 1
   - Sub item 1.1
   - Sub item 1.2
- Main item 2
   - Sub item 2.1
"""
        result = save_test_document(markdown, "list_nested.docx")
        assert not result.startswith("Error")

    def test_list_with_formatting(self):
        """Test list items with inline formatting."""
        markdown = """- **Bold item**
- *Italic item*
- Item with `code`
- Item with [link](https://example.com)
"""
        result = save_test_document(markdown, "list_formatted.docx")
        assert not result.startswith("Error")

    def test_mixed_list_types(self):
        """Test document with both ordered and unordered lists."""
        markdown = """## Shopping List

- Apples
- Bananas
- Oranges

## Steps to Follow

1. First step
2. Second step
3. Third step
"""
        result = save_test_document(markdown, "list_mixed.docx")
        assert not result.startswith("Error")


# =============================================================================
# Table Tests
# =============================================================================

class TestTables:
    """Tests for markdown table conversion."""

    def test_simple_table(self):
        """Test simple table conversion."""
        markdown = """| Name | Age | City |
|------|-----|------|
| John | 25  | NYC  |
| Jane | 30  | LA   |
"""
        result = save_test_document(markdown, "table_simple.docx")
        assert not result.startswith("Error")

    def test_table_with_formatting(self):
        """Test table with formatted cells."""
        markdown = """| Feature | Description |
|---------|-------------|
| **Bold** | This is bold |
| *Italic* | This is italic |
| `Code` | This is code |
"""
        result = save_test_document(markdown, "table_formatted.docx")
        assert not result.startswith("Error")

    def test_table_with_alignment(self):
        """Test table with column alignment markers."""
        markdown = """| Left | Center | Right |
|:-----|:------:|------:|
| L1   | C1     | R1    |
| L2   | C2     | R2    |
"""
        result = save_test_document(markdown, "table_aligned.docx")
        assert not result.startswith("Error")


# =============================================================================
# Inline Formatting Tests
# =============================================================================

class TestInlineFormatting:
    """Tests for inline markdown formatting."""

    def test_bold_text(self):
        """Test bold text conversion."""
        markdown = "This is **bold** text."
        result = save_test_document(markdown, "format_bold.docx")
        assert not result.startswith("Error")

    def test_italic_text(self):
        """Test italic text conversion."""
        markdown = "This is *italic* text."
        result = save_test_document(markdown, "format_italic.docx")
        assert not result.startswith("Error")

    def test_inline_code(self):
        """Test inline code conversion."""
        markdown = "Use the `print()` function."
        result = save_test_document(markdown, "format_code.docx")
        assert not result.startswith("Error")

    def test_hyperlink(self):
        """Test hyperlink conversion."""
        markdown = "Visit [our website](https://example.com) for more info."
        result = save_test_document(markdown, "format_link.docx")
        assert not result.startswith("Error")

    def test_mixed_formatting(self):
        """Test multiple formatting types in one paragraph."""
        markdown = "This has **bold**, *italic*, `code`, and [link](https://test.com)."
        result = save_test_document(markdown, "format_mixed.docx")
        assert not result.startswith("Error")

    def test_nested_formatting(self):
        """Test nested formatting (bold containing italic)."""
        markdown = "This is **bold with *italic* inside**."
        result = save_test_document(markdown, "format_nested.docx")
        assert not result.startswith("Error")

    def test_escaped_characters(self):
        """Test escaped markdown characters."""
        markdown = r"This has \*asterisks\* and \**double asterisks\**."
        result = save_test_document(markdown, "format_escaped.docx")
        assert not result.startswith("Error")


# =============================================================================
# Block Quote Tests
# =============================================================================

class TestBlockQuotes:
    """Tests for block quote conversion."""

    def test_simple_quote(self):
        """Test simple block quote."""
        markdown = "> This is a quoted text."
        result = save_test_document(markdown, "quote_simple.docx")
        assert not result.startswith("Error")

    def test_quote_with_formatting(self):
        """Test block quote with inline formatting."""
        markdown = "> This quote has **bold** and *italic* text."
        result = save_test_document(markdown, "quote_formatted.docx")
        assert not result.startswith("Error")


# =============================================================================
# Complex Document Tests
# =============================================================================

class TestComplexDocuments:
    """Tests for complex documents combining multiple elements."""

    def test_full_document(self):
        """Test a complete document with all elements."""
        markdown = """# Project Report

## Executive Summary

This report provides a **comprehensive analysis** of the project status.

## Key Findings

The following points summarize our findings:

- Revenue increased by **15%**
- Customer satisfaction improved to *92%*
- New features deployed successfully

## Data Overview

| Metric | Q1 | Q2 | Q3 |
|--------|----|----|-----|
| Sales | 100 | 120 | 150 |
| Users | 500 | 600 | 800 |

## Next Steps

1. Expand into new markets
2. Invest in R&D
3. Focus on customer retention

> "The best way to predict the future is to create it." - Peter Drucker

## Conclusion

Visit [our dashboard](https://example.com/dashboard) for live updates.
"""
        result = save_test_document(markdown, "complex_full_document.docx")
        assert not result.startswith("Error")

    def test_legal_contract_style(self):
        """Test legal contract style document with numbered sections."""
        markdown = """# SERVICE AGREEMENT

1. PARTIES
   - This agreement is between Company A and Company B.
   - Both parties agree to the following terms.

2. SERVICES
   - Company A will provide consulting services.
   - Services include analysis, recommendations, and implementation support.

3. PAYMENT TERMS
   - Payment is due within 30 days of invoice.
   - Late payments incur a 1.5% monthly fee.

4. CONFIDENTIALITY
   - Both parties agree to maintain confidentiality.
   - This obligation survives termination of the agreement.
"""
        result = save_test_document(markdown, "complex_contract.docx")
        assert not result.startswith("Error")

    def test_technical_documentation(self):
        """Test technical documentation style."""
        markdown = """# API Documentation

## Authentication

All API requests require authentication using an API key.

Use the `Authorization` header:

> Authorization: Bearer YOUR_API_KEY

## Endpoints

### GET /users

Returns a list of users.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| page | integer | No | Page number |
| limit | integer | No | Items per page |

### POST /users

Creates a new user.

**Request Body:**

- `name` - User's full name
- `email` - User's email address
- `role` - User's role (*admin*, *user*, or *guest*)
"""
        result = save_test_document(markdown, "complex_api_docs.docx")
        assert not result.startswith("Error")


# =============================================================================
# Edge Cases
# =============================================================================

class TestEdgeCases:
    """Tests for edge cases and special scenarios."""

    def test_empty_content(self):
        """Test with empty content."""
        markdown = ""
        result = markdown_to_word(markdown)
        assert not result.startswith("Error")

    def test_only_whitespace(self):
        """Test with only whitespace."""
        markdown = "   \n\n   \n"
        result = markdown_to_word(markdown)
        assert not result.startswith("Error")

    def test_multiple_empty_lines(self):
        """Test preservation of multiple empty lines."""
        markdown = """First paragraph.


Third paragraph (after two empty lines).
"""
        result = save_test_document(markdown, "edge_empty_lines.docx")
        assert not result.startswith("Error")

    def test_unicode_content(self):
        """Test with unicode characters."""
        markdown = """# Vícejazyčný dokument

Příliš žluťoučký kůň úpěl ďábelské ódy.

日本語テキスト

Emoji: 👋 🌍 ✨
"""
        result = save_test_document(markdown, "edge_unicode.docx")
        assert not result.startswith("Error")

    def test_long_paragraph(self):
        """Test with very long paragraph."""
        long_text = "Lorem ipsum dolor sit amet. " * 50
        markdown = f"# Long Document\n\n{long_text}"
        result = save_test_document(markdown, "edge_long_paragraph.docx")
        assert not result.startswith("Error")

    def test_special_xml_characters(self):
        """Test with characters that need XML escaping."""
        markdown = "This has < and > and & characters."
        result = save_test_document(markdown, "edge_xml_chars.docx")
        assert not result.startswith("Error")

    def test_line_breaks(self):
        """Test soft line breaks (two spaces at end)."""
        markdown = """This is line one.  
This is line two (same paragraph).  
This is line three.
"""
        result = save_test_document(markdown, "edge_line_breaks.docx")
        assert not result.startswith("Error")


# =============================================================================
# Regression Tests for helpers.py changes
# =============================================================================

class TestHelpersRegression:
    """Regression tests for helpers.py functionality used by base tool."""

    def test_parse_inline_formatting_plain(self):
        """Test parse_inline_formatting with plain text."""
        doc = Document()
        para = doc.add_paragraph()
        parse_inline_formatting("Plain text", para)
        assert para.text == "Plain text"

    def test_parse_inline_formatting_bold(self):
        """Test parse_inline_formatting with bold."""
        doc = Document()
        para = doc.add_paragraph()
        parse_inline_formatting("Text with **bold** word", para)
        assert "bold" in para.text
        bold_runs = [r for r in para.runs if r.bold]
        assert len(bold_runs) > 0

    def test_parse_inline_formatting_italic(self):
        """Test parse_inline_formatting with italic."""
        doc = Document()
        para = doc.add_paragraph()
        parse_inline_formatting("Text with *italic* word", para)
        assert "italic" in para.text
        italic_runs = [r for r in para.runs if r.italic]
        assert len(italic_runs) > 0

    def test_parse_inline_formatting_code(self):
        """Test parse_inline_formatting with inline code."""
        doc = Document()
        para = doc.add_paragraph()
        parse_inline_formatting("Use `code` here", para)
        assert "code" in para.text
        code_runs = [r for r in para.runs if r.font.name == "Courier New"]
        assert len(code_runs) > 0

    def test_parse_inline_formatting_link(self):
        """Test parse_inline_formatting with hyperlink."""
        doc = Document()
        para = doc.add_paragraph()
        parse_inline_formatting("Visit [link](https://example.com)", para)
        # Check that hyperlink element exists
        hyperlinks = para._p.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hyperlink')
        assert len(hyperlinks) > 0

    def test_parse_inline_formatting_nested(self):
        """Test parse_inline_formatting with nested formatting."""
        doc = Document()
        para = doc.add_paragraph()
        parse_inline_formatting("This is **bold with *italic* inside**", para)

        # Should have runs with both bold and italic
        bold_italic_runs = [r for r in para.runs if r.bold and r.italic]
        assert len(bold_italic_runs) > 0

    def test_parse_inline_formatting_multiple_bold(self):
        """Test parse_inline_formatting with multiple bold sections."""
        doc = Document()
        para = doc.add_paragraph()
        parse_inline_formatting("**First** and **second** bold", para)

        bold_runs = [r for r in para.runs if r.bold and r.text.strip()]
        assert len(bold_runs) >= 2


if __name__ == "__main__":
    pytest.main([__file__, "-v", "--tb=short"])

