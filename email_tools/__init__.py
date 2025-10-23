"""Email draft package exposing create_eml function.

This package provides functionality to generate EML draft files using
HTML templates that already include all CSS styling. The primary entry
point is create_eml.
"""
from .base_email_tool import create_eml  # re-export for convenience

__all__ = ["create_eml"]

