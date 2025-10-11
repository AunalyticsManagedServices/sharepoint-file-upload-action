# -*- coding: utf-8 -*-
"""
Shared utility functions for SharePoint sync operations.

This module provides common helper functions used across multiple modules.
"""

import os


def get_library_name_from_path(upload_path):
    """
    Extract library name from upload path.

    Args:
        upload_path (str): The SharePoint upload path (e.g., "Documents/folder")

    Returns:
        str: The document library name (defaults to "Documents")
    """
    library_name = "Documents"  # Default document library name
    if upload_path and "/" in upload_path:
        # If upload_path starts with a library name, use it
        path_parts = upload_path.split("/")
        if path_parts[0]:
            library_name = path_parts[0]
    return library_name


def is_debug_metadata_enabled():
    """
    Check if debug metadata mode is enabled via DEBUG_METADATA environment variable.

    This is for detailed Graph API debugging, field inspection, and metadata operations.

    Returns:
        bool: True if debug metadata mode is enabled, False otherwise
    """
    return os.environ.get('DEBUG_METADATA', 'false').lower() == 'true'


def is_debug_enabled():
    """
    Check if general debug mode is enabled via DEBUG environment variable.

    This controls individual file operation messages, folder operations, hash comparisons,
    and other verbose sync operation details. Does not affect:
    - Initial connection messages
    - File discovery/count
    - Final summary statistics
    - Rate limiting summary
    - Error messages
    - DEBUG_METADATA output (separate control)

    Returns:
        bool: True if general debug mode is enabled, False otherwise
    """
    return os.environ.get('DEBUG', 'false').lower() == 'true'
