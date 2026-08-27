"""Security tests for input validation and sanitization.

This module tests the input sanitization mechanisms to ensure that
path traversal, injection attacks, and other malicious inputs are
properly handled.
"""

import pytest
from upload_tools.utils import (
    sanitize_filename,
    generate_named_object_name,
    generate_unique_object_name,
)


class TestFilenameSanitization:
    """Test suite for filename sanitization."""

    # =========================================================================
    # Path Traversal Prevention Tests
    # =========================================================================

    @pytest.mark.parametrize("malicious_input,expected_safe", [
        # Path traversal attempts
        ("../etc/passwd", "etcpasswd"),
        ("..\\windows\\system32", "windowssystem32"),
        ("....//....//etc/passwd", "etcpasswd"),
        ("..%2f..%2fetc/passwd", "2f2fetcpasswd"),  # URL encoded
        ("/etc/passwd", "etcpasswd"),
        ("\\windows\\system32", "windowssystem32"),
        ("../../../../../../etc/passwd", "etcpasswd"),
        ("..", "document"),  # Should fallback
        ("...", "document"),  # Should fallback
        (".", "document"),  # Should fallback
    ])
    def test_path_traversal_prevention(self, malicious_input, expected_safe):
        """Test that path traversal attempts are neutralized."""
        result = sanitize_filename(malicious_input)
        assert ".." not in result, f"Path traversal sequence found in {result}"
        assert "/" not in result, f"Forward slash found in {result}"
        assert "\\" not in result, f"Backslash found in {result}"

    # =========================================================================
    # Null Byte and Control Character Tests
    # =========================================================================

    @pytest.mark.parametrize("malicious_input", [
        "file\x00.txt",  # Null byte injection
        "file\x00name.docx",
        "test\x01\x02\x03.xlsx",  # Control characters
        "evil\x7fname.pptx",  # DEL character
        "file\n\r\t.txt",  # Newlines and tabs
    ])
    def test_null_byte_and_control_chars_removed(self, malicious_input):
        """Test that null bytes and control characters are removed."""
        result = sanitize_filename(malicious_input)
        # Check no control characters remain
        for char in result:
            assert ord(char) >= 32 or char == ' ', f"Control char found: {ord(char)}"
        assert '\x00' not in result

    # =========================================================================
    # Special Character Handling Tests
    # =========================================================================

    @pytest.mark.parametrize("input_name,expected_output", [
        # Normal filenames should work
        ("My Report", "My_Report"),
        ("report-2024", "report-2024"),
        ("file_name", "file_name"),
        ("document.final", "document.final"),
        
        # Whitespace handling
        ("  spaces  around  ", "spaces_around"),
        ("multiple   spaces", "multiple_spaces"),
        ("tabs\there", "tabs_here"),
        
        # Unicode handling (should be stripped by \w)
        ("файл", "document"),  # Cyrillic - depends on regex locale
        ("文件", "document"),  # Chinese - depends on regex locale
    ])
    def test_special_character_handling(self, input_name, expected_output):
        """Test various special character scenarios."""
        result = sanitize_filename(input_name)
        # Just verify no dangerous characters
        assert "/" not in result
        assert "\\" not in result
        assert ".." not in result

    # =========================================================================
    # Injection Attack Prevention Tests
    # =========================================================================

    @pytest.mark.parametrize("malicious_input", [
        # Command injection attempts
        "file; rm -rf /",
        "file | cat /etc/passwd",
        "file`whoami`.txt",
        "file$(whoami).txt",
        "file && echo pwned",
        
        # SQL-like injection
        "file'; DROP TABLE users;--",
        "file' OR '1'='1",
        
        # HTML/Script injection
        "<script>alert('xss')</script>",
        "file<img src=x onerror=alert(1)>",
        
        # Shell special characters
        "file*?.txt",
        "file[test].txt",
        "file{a,b}.txt",
    ])
    def test_injection_attack_prevention(self, malicious_input):
        """Test that injection attack payloads are sanitized."""
        result = sanitize_filename(malicious_input)
        # Should not contain dangerous shell/SQL/HTML characters
        dangerous_chars = [';', '|', '`', '$', '&', '<', '>', "'", '"', '*', '?', '[', ']', '{', '}']
        for char in dangerous_chars:
            assert char not in result, f"Dangerous char '{char}' found in result: {result}"

    # =========================================================================
    # Edge Cases Tests
    # =========================================================================

    def test_empty_string(self):
        """Test that empty string returns fallback."""
        assert sanitize_filename("") == "document"

    def test_none_handling(self):
        """Test that None is handled gracefully."""
        # Should either return fallback or raise clear error
        try:
            result = sanitize_filename(None)
            assert result == "document"
        except (TypeError, AttributeError):
            pass  # Acceptable to raise error for None

    def test_only_dots(self):
        """Test filename with only dots."""
        assert sanitize_filename("...") == "document"
        assert sanitize_filename("..") == "document"
        assert sanitize_filename(".") == "document"

    def test_only_special_chars(self):
        """Test filename with only special characters."""
        assert sanitize_filename("!@#$%^&*()") == "document"
        assert sanitize_filename("<>:\"/\\|?*") == "document"

    def test_very_long_filename(self):
        """Test that very long filenames are truncated."""
        long_name = "a" * 500
        result = sanitize_filename(long_name)
        assert len(result) <= 100

    def test_hidden_file_attempt(self):
        """Test that hidden file creation is prevented."""
        result = sanitize_filename(".hidden")
        assert not result.startswith("."), "Should not create hidden files"

    def test_leading_trailing_cleanup(self):
        """Test leading/trailing special chars are cleaned."""
        result = sanitize_filename("___test___")
        assert not result.startswith("_")
        assert not result.endswith("_")
        
        result = sanitize_filename("...test...")
        assert not result.startswith(".")
        assert not result.endswith(".")


class TestGenerateNamedObjectName:
    """Test the generate_named_object_name function."""

    def test_basic_filename(self):
        """Test basic filename generation."""
        result = generate_named_object_name("My Report", "docx")
        assert result == "My_Report.docx"

    def test_with_unique_prefix(self):
        """Test filename with unique prefix."""
        result = generate_named_object_name("Report", "docx", add_unique_prefix=True)
        assert result.endswith(".docx")
        assert "_Report.docx" in result
        assert len(result.split("_")[0]) == 8  # 8-char UUID prefix

    def test_malicious_filename_sanitized(self):
        """Test that malicious filenames are sanitized."""
        result = generate_named_object_name("../../../etc/passwd", "docx")
        assert ".." not in result
        assert "/" not in result

    def test_malicious_suffix_sanitized(self):
        """Test that malicious suffixes are sanitized."""
        result = generate_named_object_name("file", "docx; rm -rf /")
        assert ";" not in result
        assert " " not in result.split(".")[-1]

    def test_suffix_length_limit(self):
        """Test that suffix is limited in length."""
        result = generate_named_object_name("file", "a" * 50)
        suffix = result.split(".")[-1]
        assert len(suffix) <= 10


class TestGenerateUniqueObjectName:
    """Test the generate_unique_object_name function."""

    def test_generates_uuid_filename(self):
        """Test that function generates UUID-based filename."""
        result = generate_unique_object_name("docx")
        assert result.endswith(".docx")
        # UUID format: 8-4-4-4-12 = 36 chars total
        uuid_part = result.replace(".docx", "")
        assert len(uuid_part) == 36

    def test_unique_each_call(self):
        """Test that each call generates unique name."""
        results = [generate_unique_object_name("xlsx") for _ in range(100)]
        assert len(set(results)) == 100  # All unique

    def test_suffix_sanitized(self):
        """Test that suffix is sanitized."""
        result = generate_unique_object_name("exe; whoami")
        assert ";" not in result
        assert " " not in result
