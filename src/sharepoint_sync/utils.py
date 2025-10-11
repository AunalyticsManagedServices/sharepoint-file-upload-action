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


def is_debug_mode():
    """
    Check if debug mode is enabled via DEBUG_METADATA environment variable.

    Returns:
        bool: True if debug mode is enabled, False otherwise
    """
    return os.environ.get('DEBUG_METADATA', 'false').lower() == 'true'
