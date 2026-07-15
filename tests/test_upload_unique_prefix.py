"""Tests for add_unique_prefix parameter in upload_tools.

These tests verify that the add_unique_prefix parameter correctly controls
whether a UUID prefix is added to filenames during upload.
"""

import pytest
from upload_tools.utils import generate_named_object_name


def test_generate_named_object_name_without_prefix():
    """Default behavior: no prefix (add_unique_prefix=False)."""
    result = generate_named_object_name("My Report", "docx")
    assert result == "My_Report.docx"
    # Verify no UUID prefix exists
    assert not result.startswith(("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f"))


def test_generate_named_object_name_without_prefix_explicit():
    """Explicitly pass add_unique_prefix=False."""
    result = generate_named_object_name("My Report", "docx", add_unique_prefix=False)
    assert result == "My_Report.docx"


def test_generate_named_object_name_with_prefix():
    """With add_unique_prefix=True: 8-char UUID prefix."""
    result = generate_named_object_name("My Report", "docx", add_unique_prefix=True)
    assert result.endswith("_My_Report.docx")
    # Extract prefix (everything before first underscore)
    prefix = result.split("_")[0]
    assert len(prefix) == 8
    # Verify prefix is hexadecimal
    assert all(c in "0123456789abcdef" for c in prefix)


def test_generate_named_object_name_with_prefix_different_files():
    """Verify different UUIDs for different calls with same filename."""
    result1 = generate_named_object_name("document", "xlsx", add_unique_prefix=True)
    result2 = generate_named_object_name("document", "xlsx", add_unique_prefix=True)
    
    # Both should end with same filename
    assert result1.endswith("_document.xlsx")
    assert result2.endswith("_document.xlsx")
    
    # But prefixes should be different (different UUIDs)
    prefix1 = result1.split("_")[0]
    prefix2 = result2.split("_")[0]
    assert prefix1 != prefix2


def test_generate_named_object_name_sanitization():
    """Sanitization still works with and without prefix."""
    # Without prefix
    result = generate_named_object_name("Report: Q1 2024!", "xlsx", add_unique_prefix=False)
    assert result == "Report_Q1_2024.xlsx"
    
    # With prefix
    result_with_prefix = generate_named_object_name("Report: Q1 2024!", "xlsx", add_unique_prefix=True)
    assert result_with_prefix.endswith("_Report_Q1_2024.xlsx")
    prefix = result_with_prefix.split("_")[0]
    assert len(prefix) == 8


def test_generate_named_object_name_special_characters():
    """Test sanitization of special characters."""
    # Without prefix
    result = generate_named_object_name("My@Document#2024", "docx", add_unique_prefix=False)
    assert result == "MyDocument2024.docx"
    
    # With prefix
    result_with_prefix = generate_named_object_name("My@Document#2024", "docx", add_unique_prefix=True)
    assert result_with_prefix.endswith("_MyDocument2024.docx")


def test_generate_named_object_name_different_extensions():
    """Test with different file extensions."""
    extensions = ["docx", "xlsx", "pptx", "eml", "xml"]
    
    for ext in extensions:
        # Without prefix
        result = generate_named_object_name("test", ext, add_unique_prefix=False)
        assert result == f"test.{ext}"
        
        # With prefix
        result_with_prefix = generate_named_object_name("test", ext, add_unique_prefix=True)
        assert result_with_prefix.endswith(f"_test.{ext}")
        prefix = result_with_prefix.split("_")[0]
        assert len(prefix) == 8


def test_generate_named_object_name_empty_string():
    """Test with empty filename (should fallback to 'document')."""
    # Without prefix
    result = generate_named_object_name("", "docx", add_unique_prefix=False)
    assert result == "document.docx"
    
    # With prefix
    result_with_prefix = generate_named_object_name("", "docx", add_unique_prefix=True)
    assert result_with_prefix.endswith("_document.docx")


def test_generate_named_object_name_whitespace_only():
    """Test with whitespace-only filename (should fallback to 'document')."""
    # Without prefix
    result = generate_named_object_name("   ", "xlsx", add_unique_prefix=False)
    assert result == "document.xlsx"
    
    # With prefix
    result_with_prefix = generate_named_object_name("   ", "xlsx", add_unique_prefix=True)
    assert result_with_prefix.endswith("_document.xlsx")


def test_generate_named_object_name_long_filename():
    """Test with long filename (should truncate to 100 chars)."""
    long_name = "a" * 150
    
    # Without prefix
    result = generate_named_object_name(long_name, "docx", add_unique_prefix=False)
    assert result == f"{long_name[:100]}.docx"
    
    # With prefix (truncation happens after prefix)
    result_with_prefix = generate_named_object_name(long_name, "docx", add_unique_prefix=True)
    assert result_with_prefix.endswith(f"_{long_name[:100]}.docx")
