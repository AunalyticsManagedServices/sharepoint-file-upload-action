# -*- coding: utf-8 -*-
"""
SharePoint File Upload Sync Package
====================================

This package provides modular components for syncing files from GitHub
repositories to SharePoint document libraries with intelligent change detection,
hash-based comparison, and Markdown-to-HTML conversion.

Modules:
--------
- config: Configuration and argument parsing
- auth: Microsoft authentication
- graph_api: Microsoft Graph API operations
- file_handler: File operations (hashing, sanitization, comparison)
- uploader: Upload operations and folder management
- markdown_converter: Markdown to HTML conversion with Mermaid diagrams
- monitoring: Rate limiting monitoring and statistics tracking
- utils: Shared utility functions

Usage Example:
-------------
    from sharepoint_sync.config import parse_config
    from sharepoint_sync.graph_api import create_graph_client

    cfg = parse_config()
    client = create_graph_client(
        cfg.tenant_id,
        cfg.client_id,
        cfg.client_secret,
        cfg.login_endpoint,
        cfg.graph_endpoint
    )
"""

__version__ = "2.0.0"
__author__ = "Mark Newton"

# Main exports for convenience
from .config import parse_config, Config
from .auth import acquire_token
from .graph_api import (
    create_graph_client,
    check_and_create_filehash_column,
    comprehensive_column_verification,
    verify_column_for_filehash_operations,
    test_column_accessibility,
    list_files_in_folder_recursive,
    delete_file_from_sharepoint
)
from .monitoring import upload_stats, rate_monitor, print_rate_limiting_summary
from .file_handler import (
    calculate_file_hash,
    sanitize_sharepoint_name,
    check_file_needs_update
)
from .uploader import upload_file, upload_file_with_structure, ensure_folder_exists
from .markdown_converter import convert_markdown_to_html
from .utils import get_library_name_from_path, is_debug_metadata_enabled, is_debug_enabled

__all__ = [
    # Configuration
    'parse_config',
    'Config',
    # Authentication
    'acquire_token',
    # Graph API
    'create_graph_client',
    'check_and_create_filehash_column',
    'comprehensive_column_verification',
    'verify_column_for_filehash_operations',
    'test_column_accessibility',
    'list_files_in_folder_recursive',
    'delete_file_from_sharepoint',
    # File Operations
    'calculate_file_hash',
    'sanitize_sharepoint_name',
    'check_file_needs_update',
    # Upload Operations
    'upload_file',
    'upload_file_with_structure',
    'ensure_folder_exists',
    # Markdown
    'convert_markdown_to_html',
    # Monitoring
    'upload_stats',
    'rate_monitor',
    'print_rate_limiting_summary',
    # Utilities
    'get_library_name_from_path',
    'is_debug_metadata_enabled',
    'is_debug_enabled',
]
