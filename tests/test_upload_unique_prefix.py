"""Tests for add_unique_prefix parameter in upload_tools.

These tests verify that the add_unique_prefix parameter correctly controls
whether a UUID prefix is added to filenames during upload.
"""

import pytest
from unittest.mock import patch, MagicMock
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


# =============================================================================
# Tests for strategy-based default behavior
# =============================================================================

class TestStrategyBasedDefaults:
    """Tests for add_unique_prefix strategy-based default behavior.
    
    These tests verify that the default value of add_unique_prefix depends on
    the storage strategy:
    - Traditional backends (LOCAL, S3, GCS, AZURE, MINIO): defaults to True
    - LIBRECHAT: defaults to False (LibreChat handles its own UUID prefix)
    """

    def test_upload_file_defaults_to_true_for_local_strategy(self):
        """For LOCAL strategy, add_unique_prefix should default to True."""
        from config import StorageStrategy
        
        with patch('upload_tools.main.UPLOAD_STRATEGY', StorageStrategy.LOCAL), \
             patch('upload_tools.main.upload_to_local_folder') as mock_upload:
            mock_upload.return_value = "http://example.com/file.docx"
            
            from upload_tools.main import upload_file
            from io import BytesIO
            
            file_obj = BytesIO(b"test content")
            # Don't pass add_unique_prefix - should default to True for LOCAL
            result = upload_file(file_obj, "docx", filename="test_file")
            
            # Check that the object_name passed to upload has a UUID prefix
            call_args = mock_upload.call_args
            object_name = call_args[0][1]  # Second positional arg is object_name
            assert "_test_file.docx" in object_name, f"Expected UUID prefix, got: {object_name}"
            # Verify the prefix is 8 hex characters
            prefix = object_name.split("_")[0]
            assert len(prefix) == 8, f"Expected 8-char prefix, got: {prefix}"
            assert all(c in "0123456789abcdef" for c in prefix), f"Prefix not hex: {prefix}"

    def test_upload_file_respects_explicit_false_for_local_strategy(self):
        """For LOCAL strategy, explicit add_unique_prefix=False should be respected."""
        from config import StorageStrategy
        
        with patch('upload_tools.main.UPLOAD_STRATEGY', StorageStrategy.LOCAL), \
             patch('upload_tools.main.upload_to_local_folder') as mock_upload:
            mock_upload.return_value = "http://example.com/file.docx"
            
            from upload_tools.main import upload_file
            from io import BytesIO
            
            file_obj = BytesIO(b"test content")
            # Explicitly pass add_unique_prefix=False
            result = upload_file(file_obj, "docx", filename="test_file", add_unique_prefix=False)
            
            # Check that the object_name has NO UUID prefix
            call_args = mock_upload.call_args
            object_name = call_args[0][1]
            assert object_name == "test_file.docx", f"Expected no prefix, got: {object_name}"

    @pytest.mark.asyncio
    async def test_upload_file_async_defaults_to_false_for_librechat_strategy(self):
        """For LIBRECHAT strategy, add_unique_prefix should default to False."""
        from config import StorageStrategy
        
        with patch('upload_tools.main.UPLOAD_STRATEGY', StorageStrategy.LIBRECHAT), \
             patch('upload_tools.backends.librechat.upload_to_librechat') as mock_upload:
            mock_upload.return_value = {"file_id": "123", "filename": "test_file.docx"}
            
            from upload_tools.main import upload_file_async
            from io import BytesIO
            
            file_obj = BytesIO(b"test content")
            user_context = {"user_id": "test_user"}
            
            # Don't pass add_unique_prefix - should default to False for LIBRECHAT
            result = await upload_file_async(
                file_obj, "docx", 
                filename="test_file", 
                user_context=user_context
            )
            
            # Check that the object_name passed to upload has NO UUID prefix
            call_args = mock_upload.call_args
            object_name = call_args[0][1]  # Second positional arg is object_name
            assert object_name == "test_file.docx", f"Expected no prefix for LIBRECHAT, got: {object_name}"

    @pytest.mark.asyncio
    async def test_upload_file_async_respects_explicit_true_for_librechat_strategy(self):
        """For LIBRECHAT strategy, explicit add_unique_prefix=True should be respected."""
        from config import StorageStrategy
        
        with patch('upload_tools.main.UPLOAD_STRATEGY', StorageStrategy.LIBRECHAT), \
             patch('upload_tools.backends.librechat.upload_to_librechat') as mock_upload:
            mock_upload.return_value = {"file_id": "123", "filename": "test_file.docx"}
            
            from upload_tools.main import upload_file_async
            from io import BytesIO
            
            file_obj = BytesIO(b"test content")
            user_context = {"user_id": "test_user"}
            
            # Explicitly pass add_unique_prefix=True
            result = await upload_file_async(
                file_obj, "docx", 
                filename="test_file", 
                user_context=user_context,
                add_unique_prefix=True
            )
            
            # Check that the object_name has a UUID prefix
            call_args = mock_upload.call_args
            object_name = call_args[0][1]
            assert "_test_file.docx" in object_name, f"Expected UUID prefix, got: {object_name}"

    @pytest.mark.asyncio
    async def test_upload_file_async_defaults_to_true_for_s3_strategy(self):
        """For S3 strategy, add_unique_prefix should default to True."""
        from config import StorageStrategy
        
        with patch('upload_tools.main.UPLOAD_STRATEGY', StorageStrategy.S3), \
             patch('upload_tools.main.upload_to_s3') as mock_upload, \
             patch('upload_tools.main.cfg') as mock_cfg:
            mock_upload.return_value = "https://s3.example.com/file.docx"
            mock_cfg.storage.s3 = MagicMock()
            
            from upload_tools.main import upload_file_async
            from io import BytesIO
            
            file_obj = BytesIO(b"test content")
            
            # Don't pass add_unique_prefix - should default to True for S3
            result = await upload_file_async(file_obj, "docx", filename="test_file")
            
            # Check that the object_name has a UUID prefix
            call_args = mock_upload.call_args
            object_name = call_args[0][1]
            assert "_test_file.docx" in object_name, f"Expected UUID prefix for S3, got: {object_name}"


class TestToolHandlerIntegration:
    """Integration tests for add_unique_prefix through actual tool handlers.
    
    These tests verify that the strategy-based default resolution works
    end-to-end through the MCP tool handlers, not just the upload functions.
    """

    @pytest.mark.asyncio
    async def test_tool_handler_defaults_to_uuid_prefix_for_local_strategy(self):
        """For LOCAL strategy, tool handlers should add UUID prefix by default.
        
        This tests the full flow: tool handler → upload_and_format_response → 
        upload_file → generate_named_object_name, verifying that the None default
        resolves to True for traditional backends.
        """
        from config import StorageStrategy
        
        # Mock at the upload_tools level to capture the actual call
        with patch('librechat_integration.is_librechat_strategy', return_value=False), \
             patch('upload_tools.main.UPLOAD_STRATEGY', StorageStrategy.LOCAL), \
             patch('upload_tools.main.upload_to_local_folder') as mock_local_upload:
            
            mock_local_upload.return_value = "http://example.com/uuid_test_document.docx"
            
            from librechat_integration import upload_and_format_response
            from io import BytesIO
            
            file_obj = BytesIO(b"test content")
            user_context = {"user_id": None}  # Not LibreChat, user_id not required
            
            # Call upload_and_format_response with add_unique_prefix=None (simulating tool handler default)
            result = await upload_and_format_response(
                file_obj,
                "docx",
                "test_document",
                user_context,
                "Test document created.",
                add_unique_prefix=None,  # This is what tool handlers pass by default now
            )
            
            # Verify upload_to_local_folder was called with a UUID-prefixed filename
            mock_local_upload.assert_called_once()
            call_args = mock_local_upload.call_args
            object_name = call_args[0][1]  # Second positional arg is object_name
            
            # The key assertion: filename should have UUID prefix because None resolved to True
            assert "_test_document.docx" in object_name, \
                f"Expected UUID prefix when add_unique_prefix=None for LOCAL strategy, got: {object_name}"
            prefix = object_name.split("_")[0]
            assert len(prefix) == 8, f"Expected 8-char UUID prefix, got: {prefix}"
            assert all(c in "0123456789abcdef" for c in prefix), f"Prefix not hex: {prefix}"

    @pytest.mark.asyncio
    async def test_upload_file_resolves_none_to_true_for_local(self):
        """Verify upload_file resolves None to True for LOCAL strategy."""
        from config import StorageStrategy
        
        with patch('upload_tools.main.UPLOAD_STRATEGY', StorageStrategy.LOCAL), \
             patch('upload_tools.main.upload_to_local_folder') as mock_upload:
            mock_upload.return_value = "http://example.com/uuid_test_file.docx"
            
            from upload_tools.main import upload_file
            from io import BytesIO
            
            file_obj = BytesIO(b"test content")
            
            # Call with add_unique_prefix=None (default for tool handlers)
            result = upload_file(file_obj, "docx", filename="test_file", add_unique_prefix=None)
            
            # Check that the object_name passed to upload has a UUID prefix
            call_args = mock_upload.call_args
            object_name = call_args[0][1]  # Second positional arg is object_name
            assert "_test_file.docx" in object_name, \
                f"Expected UUID prefix when add_unique_prefix=None for LOCAL, got: {object_name}"
            # Verify the prefix is 8 hex characters
            prefix = object_name.split("_")[0]
            assert len(prefix) == 8, f"Expected 8-char prefix, got: {prefix}"
            assert all(c in "0123456789abcdef" for c in prefix), f"Prefix not hex: {prefix}"


class TestToolSignatureDefaults:
    """Tests that verify the actual FastMCP tool handler signatures have correct defaults.
    
    These tests use AST parsing to inspect main.py source code and verify that the
    add_unique_prefix parameter defaults to None (not False) in all tool handler
    signatures. This catches regressions where someone changes the parameter default
    from None back to False.
    
    Using AST parsing avoids importing main.py (which requires openpyxl, python-docx,
    etc.) while still verifying the actual source code.
    """

    @staticmethod
    def _get_function_param_default(source_code: str, func_name: str, param_name: str):
        """Extract the default value of a parameter from a function definition using AST.
        
        Returns:
            The default value as a Python object, or a special marker if no default.
        """
        import ast
        
        tree = ast.parse(source_code)
        
        for node in ast.walk(tree):
            if isinstance(node, ast.AsyncFunctionDef) and node.name == func_name:
                # Get all arguments and their defaults
                args = node.args
                # args.args contains positional args, args.defaults contains their defaults
                # defaults are right-aligned: if there are 5 args and 3 defaults,
                # the first 2 args have no default
                
                # Also check kwonlyargs and kw_defaults
                all_args = args.args + args.kwonlyargs
                all_defaults = (
                    [None] * (len(args.args) - len(args.defaults)) + 
                    list(args.defaults) + 
                    list(args.kw_defaults)
                )
                
                for arg, default in zip(all_args, all_defaults):
                    if arg.arg == param_name:
                        if default is None:
                            return "NO_DEFAULT"
                        elif isinstance(default, ast.Constant):
                            return default.value
                        elif isinstance(default, ast.NameConstant):  # Python 3.7
                            return default.value
                        elif isinstance(default, ast.Name):
                            return default.id  # e.g., 'None' as a name
                        else:
                            return f"COMPLEX_DEFAULT: {ast.dump(default)}"
                
                return "PARAM_NOT_FOUND"
        
        return "FUNCTION_NOT_FOUND"

    @staticmethod
    def _get_main_py_source():
        """Read main.py source code."""
        from pathlib import Path
        main_py = Path(__file__).parent.parent / "main.py"
        return main_py.read_text()

    def test_create_excel_document_signature_defaults_to_none(self):
        """Verify create_excel_document's add_unique_prefix defaults to None (not False).
        
        If the signature is changed to default=False, this test will fail.
        """
        source = self._get_main_py_source()
        default = self._get_function_param_default(source, "create_excel_document", "add_unique_prefix")
        
        assert default is None, \
            f"create_excel_document's add_unique_prefix should default to None, got: {default!r}"

    def test_create_word_document_signature_defaults_to_none(self):
        """Verify create_word_document's add_unique_prefix defaults to None (not False)."""
        source = self._get_main_py_source()
        default = self._get_function_param_default(source, "create_word_document", "add_unique_prefix")
        
        assert default is None, \
            f"create_word_document's add_unique_prefix should default to None, got: {default!r}"

    def test_create_powerpoint_presentation_signature_defaults_to_none(self):
        """Verify create_powerpoint_presentation's add_unique_prefix defaults to None (not False)."""
        source = self._get_main_py_source()
        default = self._get_function_param_default(source, "create_powerpoint_presentation", "add_unique_prefix")
        
        assert default is None, \
            f"create_powerpoint_presentation's add_unique_prefix should default to None, got: {default!r}"

    def test_create_email_draft_signature_defaults_to_none(self):
        """Verify create_email_draft's add_unique_prefix defaults to None (not False)."""
        source = self._get_main_py_source()
        default = self._get_function_param_default(source, "create_email_draft", "add_unique_prefix")
        
        assert default is None, \
            f"create_email_draft's add_unique_prefix should default to None, got: {default!r}"

    def test_create_xml_document_signature_defaults_to_none(self):
        """Verify create_xml_document's add_unique_prefix defaults to None (not False)."""
        source = self._get_main_py_source()
        default = self._get_function_param_default(source, "create_xml_document", "add_unique_prefix")
        
        assert default is None, \
            f"create_xml_document's add_unique_prefix should default to None, got: {default!r}"
