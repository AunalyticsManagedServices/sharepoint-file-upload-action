# -*- coding: utf-8 -*-
"""
File handling operations for SharePoint sync.

This module provides functions for file sanitization, hashing, comparison, and exclusion.
"""

import os
import xxhash
import fnmatch


def sanitize_sharepoint_name(name, is_folder=False):
    r"""
    Sanitize file/folder names to be compatible with SharePoint/OneDrive.

    SharePoint/OneDrive has strict naming rules:
    - Cannot contain: # % & * : < > ? / \ | " { } ~
    - Cannot start with: ~ $
    - Cannot end with: . (period)
    - Cannot be reserved names: CON, PRN, AUX, NUL, COM1-9, LPT1-9
    - Maximum length: 400 characters for full path, 255 for file/folder name

    Args:
        name (str): Original file or folder name
        is_folder (bool): Whether this is a folder name

    Returns:
        str: Sanitized name safe for SharePoint
    """
    if not name:
        return name

    # Map of illegal characters to safe replacements
    # Using Unicode similar characters that are visually similar but allowed
    char_replacements = {
        '#': '＃',    # Fullwidth number sign
        '%': '％',    # Fullwidth percent sign
        '&': '＆',    # Fullwidth ampersand
        '*': '＊',    # Fullwidth asterisk
        ':': '：',    # Fullwidth colon
        '<': '＜',    # Fullwidth less-than
        '>': '＞',    # Fullwidth greater-than
        '?': '？',    # Fullwidth question mark
        '/': '／',    # Fullwidth solidus
        '\\': '＼',   # Fullwidth reverse solidus
        '|': '｜',    # Fullwidth vertical line
        '"': '＂',    # Fullwidth quotation mark
        '{': '｛',    # Fullwidth left curly bracket
        '}': '｝',    # Fullwidth right curly bracket
        '~': '～',    # Fullwidth tilde
    }

    # Start with original name
    sanitized = name

    # Replace illegal characters
    for char, replacement in char_replacements.items():
        sanitized = sanitized.replace(char, replacement)

    # Remove leading ~ or $ characters
    while sanitized and sanitized[0] in ['~', '$', '～']:
        sanitized = sanitized[1:]

    # Remove trailing periods and spaces
    sanitized = sanitized.rstrip('. ')

    # Check for reserved names (Windows legacy)
    reserved_names = [
        'CON', 'PRN', 'AUX', 'NUL',
        'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9',
        'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9'
    ]

    # Check if name (without extension) is reserved
    name_without_ext = sanitized.split('.')[0] if not is_folder else sanitized
    if name_without_ext.upper() in reserved_names:
        sanitized = f"_{sanitized}"  # Prefix with underscore to make it safe

    # Ensure name isn't empty after sanitization
    if not sanitized:
        sanitized = "_unnamed"

    # Truncate if too long (SharePoint limit is 255 chars for file/folder name)
    if len(sanitized) > 255:
        # If it's a file, preserve the extension
        if not is_folder and '.' in name:
            ext = name.split('.')[-1]
            base_max_len = 255 - len(ext) - 1  # -1 for the dot
            base = sanitized[:base_max_len]
            sanitized = f"{base}.{ext}"
        else:
            sanitized = sanitized[:255]

    # Log if name was changed
    if sanitized != name:
        print(f"[!] Sanitized name: '{name}' -> '{sanitized}'")

    return sanitized

def sanitize_path_components(path):
    """
    Sanitize all components of a file path for SharePoint compatibility.

    Args:
        path (str): Full path with possibly multiple directory levels

    Returns:
        str: Sanitized path with all components made SharePoint-safe
    """
    # Split path into components
    path = path.replace('\\', '/')
    components = path.split('/')

    # Sanitize each component
    sanitized_components = []
    for i, component in enumerate(components):
        if component:  # Skip empty components
            # Last component might be a file, others are folders
            is_folder = (i < len(components) - 1) or not ('.' in component)
            sanitized = sanitize_sharepoint_name(component, is_folder)
            sanitized_components.append(sanitized)

    # Rejoin path
    return '/'.join(sanitized_components)


def get_optimal_chunk_size(file_size):
    """
    Calculate optimal chunk size based on file size for efficient hashing.

    Larger files benefit from larger chunks to reduce I/O overhead,
    while smaller files use smaller chunks to avoid memory waste.

    Args:
        file_size (int): Size of the file in bytes

    Returns:
        int: Optimal chunk size in bytes for reading the file
    """
    if file_size < 1 * 1024 * 1024:  # < 1MB
        return 64 * 1024  # 64KB chunks - small files, minimal memory
    elif file_size < 10 * 1024 * 1024:  # < 10MB
        return 256 * 1024  # 256KB chunks - balance memory/speed
    elif file_size < 100 * 1024 * 1024:  # < 100MB
        return 1 * 1024 * 1024  # 1MB chunks - larger reads for efficiency
    elif file_size < 1024 * 1024 * 1024:  # < 1GB
        return 4 * 1024 * 1024  # 4MB chunks - maximize throughput
    else:  # >= 1GB
        return 8 * 1024 * 1024  # 8MB chunks - optimal for very large files


def calculate_file_hash(file_path):
    """
    Calculate xxHash128 for a file using dynamic chunk sizing.

    xxHash128 is a non-cryptographic hash that's 10-20x faster than SHA-256
    while still providing excellent avalanche properties and collision resistance
    for file deduplication purposes.

    Args:
        file_path (str): Path to the file to hash

    Returns:
        str: Hexadecimal string representation of the xxHash128 (32 characters)

    Note:
        The hash is deterministic - same file always produces same hash
        regardless of when/where it's calculated (no timestamps involved).
    """
    try:
        file_size = os.path.getsize(file_path)
        chunk_size = get_optimal_chunk_size(file_size)

        # Use xxh128 (alias for xxh3_128) for maximum speed on modern CPUs
        hasher = xxhash.xxh128()

        with open(file_path, 'rb') as f:
            while chunk := f.read(chunk_size):
                hasher.update(chunk)

        return hasher.hexdigest()
    except Exception as e:
        print(f"[!] Error calculating hash for {file_path}: {e}")
        return None


def should_exclude_path(path, exclude_patterns):
    """
    Check if a file or directory path should be excluded based on exclusion patterns.

    This function provides cross-platform exclusion filtering using fnmatch for
    pattern matching. It checks both the full path and individual path components
    (for directory exclusions like '__pycache__' or 'node_modules').

    Args:
        path (str): File or directory path to check (can be absolute or relative)
        exclude_patterns (list): List of exclusion patterns (e.g., ['*.tmp', '*.log', '__pycache__'])

    Returns:
        bool: True if path should be excluded, False otherwise

    Pattern Matching:
        - Exact filename match: '__pycache__', '.git', 'node_modules'
        - Wildcard patterns: '*.tmp', '*.log', '*.pyc'
        - Extension only: 'tmp', 'log' (automatically converts to '*.tmp', '*.log')

    Cross-Platform Compatibility:
        - Works with both forward slashes (/) and backslashes (\\)
        - Normalizes paths for consistent matching on Windows and Linux
        - Case-sensitive on Linux, case-insensitive on Windows

    Examples:
        >>> should_exclude_path('file.tmp', ['*.tmp'])
        True
        >>> should_exclude_path('src/__pycache__/module.pyc', ['__pycache__'])
        True
        >>> should_exclude_path('docs/report.pdf', ['*.tmp', '*.log'])
        False
    """
    if not exclude_patterns:
        return False

    # Normalize path separators for cross-platform compatibility
    # Convert backslashes to forward slashes for consistent handling
    normalized_path = path.replace('\\', '/')

    # Get the basename (filename or directory name)
    basename = os.path.basename(normalized_path)

    # Split path into components for directory matching
    # This allows matching directory names anywhere in the path
    path_components = normalized_path.split('/')

    for pattern in exclude_patterns:
        # Match against basename (most common case)
        # This handles patterns like '*.tmp', '__pycache__', 'file.log'
        if fnmatch.fnmatch(basename, pattern):
            return True

        # If pattern doesn't contain wildcards, check if it matches any path component
        # This allows excluding directories like '__pycache__' or 'node_modules' anywhere in path
        if '*' not in pattern and '?' not in pattern and '[' not in pattern:
            if pattern in path_components:
                return True

        # Check if pattern matches full path (for more specific exclusions)
        if fnmatch.fnmatch(normalized_path, pattern):
            return True

        # Auto-add wildcard for extension-only patterns (e.g., 'tmp' -> '*.tmp')
        if not pattern.startswith('*') and not pattern.startswith('.'):
            wildcard_pattern = f'*.{pattern}'
            if fnmatch.fnmatch(basename, wildcard_pattern):
                return True

    return False


def check_file_needs_update(drive, local_path, file_name, site_url, list_name, filehash_column_available,
                            tenant_id=None, client_id=None, client_secret=None, login_endpoint=None,
                            graph_endpoint=None, upload_stats_dict=None):
    """
    Check if a file in SharePoint needs to be updated by comparing hash or size.

    This function implements efficient file comparison to avoid unnecessary uploads.
    Files are compared using:
    1. FileHash (xxHash128) if column exists - most reliable
    2. Size comparison as fallback - works without custom columns

    Args:
        drive (DriveItem): The folder to check for existing file
        local_path (str): Path to the local file
        file_name (str): Name of the file to check
        site_url (str): SharePoint site URL (e.g., 'company.sharepoint.com')
        list_name (str): SharePoint library name
        filehash_column_available (bool): Whether FileHash column exists
        tenant_id (str, optional): Azure AD tenant ID for REST API calls
        client_id (str, optional): Azure AD client ID for REST API calls
        client_secret (str, optional): Azure AD client secret for REST API calls
        login_endpoint (str, optional): Azure AD login endpoint for REST API calls
        graph_endpoint (str, optional): Microsoft Graph API endpoint for REST API calls
        upload_stats_dict (dict, optional): Upload statistics dictionary to update

    Returns:
        tuple: (needs_update: bool, exists: bool, remote_file: DriveItem or None, local_hash: str or None)
            - needs_update: True if file should be uploaded
            - exists: True if file exists in SharePoint
            - remote_file: The existing SharePoint file object (if exists)
            - local_hash: The calculated hash of the local file (if computed)

    Example:
        needs_update, exists, remote, hash_val = check_file_needs_update(
            folder, "/path/to/file.pdf", "file.pdf", "site.sharepoint.com", "Documents", True,
            tenant_id, client_id, client_secret, login_endpoint, graph_endpoint
        )
        if not needs_update:
            print("File is up to date, skipping")
    """

    # Sanitize the file name to match what would be stored in SharePoint
    sanitized_name = sanitize_sharepoint_name(file_name, is_folder=False)

    # Calculate local file hash upfront for efficiency
    local_hash = calculate_file_hash(local_path)
    if local_hash:
        print(f"[#] Local hash: {local_hash[:8]}... for {sanitized_name}")

    # Get local file information
    local_size = os.path.getsize(local_path)

    # Get debug flag (used throughout function)
    debug_metadata = os.environ.get('DEBUG_METADATA', 'false').lower() == 'true'

    # Debug: Show what we're checking
    print(f"[?] Checking if file exists in SharePoint: {sanitized_name}")

    try:
        # Get all children in the folder to find our file
        # This is more reliable than get_by_path for some SharePoint configurations
        # Request all needed properties upfront to avoid secondary queries
        children = drive.children.get().select(["name", "file", "folder", "size", "id"]).execute_query()

        existing_file = None
        for child in children:
            # Use getattr for safe attribute access
            child_name = getattr(child, 'name', None)
            if child_name and child_name == sanitized_name:
                existing_file = child
                break

        if existing_file is None:
            # File not found in folder
            print(f"[+] New file to upload: {sanitized_name}")
            return True, False, None, local_hash

        # First check if this is a file or folder
        # In SharePoint, items have both 'file' and 'folder' properties, but only one is populated
        # We need to check which one has actual content, not just if it exists

        # Try to determine if this is a folder
        is_folder = False
        is_file = False

        # Check if folder property exists and has meaningful content
        if hasattr(existing_file, 'folder'):
            folder_prop = getattr(existing_file, 'folder', None)
            # SharePoint may return empty objects {} instead of None
            # Check if it's not None AND has actual content (like childCount)
            if folder_prop is not None:
                # For folders, the folder property should have attributes like childCount
                # Check for multiple folder-specific attributes to be certain
                if (hasattr(folder_prop, 'childCount') and folder_prop.childCount is not None) or \
                   (hasattr(folder_prop, 'child_count') and folder_prop.child_count is not None):
                    is_folder = True
                # Don't assume it's a folder just because the property exists
                # Empty objects should not count as folders

        # Check if file property exists and has meaningful content
        if hasattr(existing_file, 'file'):
            file_prop = getattr(existing_file, 'file', None)
            # Same check for files - avoid empty objects
            if file_prop is not None:
                # Files should have properties like mimeType or hashes
                # Check if it's not just an empty object
                if hasattr(file_prop, 'mimeType') or hasattr(file_prop, 'mime_type') or hasattr(file_prop, 'hashes'):
                    is_file = True
                # Don't assume it's a file just because the property exists
                # Empty objects should not count as files

        # If we can't determine from the properties, try a different approach
        # Check if the item has a size attribute (files have size, folders typically don't)
        if not is_folder and not is_file:
            # If we have size already, it's likely a file
            if hasattr(existing_file, 'size') and existing_file.size is not None and existing_file.size > 0:
                is_file = True
            # If neither property is populated, default to treating as file
            # (better to attempt upload than to skip)
            else:
                print(f"[?] Cannot determine type for {sanitized_name}, treating as file")
                is_file = True

        # Log only if there's true ambiguity (both detected as populated)
        if is_folder and is_file:
            print(f"[!] Warning: Item {sanitized_name} appears to be both file and folder, treating as file")
            # Prefer file treatment to allow upload attempt
            is_folder = False

        if is_folder:
            # It's definitely a folder
            print(f"[!] Conflict: Folder exists with same name as file: {sanitized_name}")
            return True, False, None, local_hash

        # Treat as file if it has file property or if we can't determine type
        # (Better to try uploading than to skip)

        # Only fetch size/date properties for files, not folders
        # Skip if we determined this is a folder
        if is_folder and not is_file:
            # Don't try to get file metadata for folders
            return True, False, None, local_hash

        # First, try to get the FileHash property if it exists using direct REST API
        hash_comparison_available = False

        # Only attempt FileHash comparison if we have the necessary credentials
        if filehash_column_available and all([tenant_id, client_id, client_secret, login_endpoint, graph_endpoint]):
            try:
                # Use direct Graph API REST calls to get SharePoint list item with custom columns
                # First, try to import the function
                try:
                    from .graph_api import get_sharepoint_list_item_by_filename
                except ImportError:
                    # graph_api module not yet available
                    if debug_metadata:
                        print(f"[DEBUG] graph_api module not yet imported")
                    get_sharepoint_list_item_by_filename = None

                # Call the function if it was successfully imported
                if get_sharepoint_list_item_by_filename is not None:
                    list_item_data = get_sharepoint_list_item_by_filename(
                        site_url, list_name, sanitized_name,
                        tenant_id, client_id, client_secret, login_endpoint, graph_endpoint
                    )
                else:
                    list_item_data = None

                if list_item_data and 'fields' in list_item_data:
                    fields = list_item_data['fields']

                    if debug_metadata:
                        print(f"[DEBUG] Retrieving FileHash for {sanitized_name}")
                        print(f"[DEBUG] REST API list item data: {type(list_item_data)}")
                        print(f"[DEBUG] fields data: {type(fields)}")
                        print(f"[DEBUG] Available field properties: {list(fields.keys())}")
                        print(f"[DEBUG] FileHash in properties: {'FileHash' in fields}")
                        if 'FileHash' in fields:
                            print(f"[DEBUG] FileHash value: {fields.get('FileHash')}")

                    # Access the FileHash custom column from the fields
                    remote_hash = fields.get('FileHash')

                    if remote_hash:
                        hash_comparison_available = True
                        print(f"[#] Remote hash: {remote_hash[:8]}... for {sanitized_name}")

                        # Compare hashes - this is the most reliable comparison
                        if local_hash and local_hash == remote_hash:
                            print(f"[=] File unchanged (hash match): {sanitized_name}")
                            if upload_stats_dict:
                                upload_stats_dict['skipped_files'] += 1
                                upload_stats_dict['bytes_skipped'] += local_size
                            return False, True, existing_file, local_hash
                        elif local_hash:
                            print(f"[*] File changed (hash mismatch): {sanitized_name}")
                            return True, True, existing_file, local_hash
                    elif debug_metadata:
                        print(f"[DEBUG] FileHash not found in list item fields")
                elif debug_metadata:
                    print(f"[DEBUG] Could not retrieve list item data for {sanitized_name}")

            except Exception as hash_error:
                # FileHash column might not exist, or we can't access it
                print(f"[!] Could not retrieve FileHash via REST API, falling back to size comparison: {str(hash_error)[:100]}")
                hash_comparison_available = False
        else:
            # FileHash column not available or missing credentials
            if debug_metadata:
                if not filehash_column_available:
                    print(f"[DEBUG] FileHash column not available, using size comparison")
                else:
                    print(f"[DEBUG] Missing credentials for FileHash retrieval, using size comparison")

        # If hash comparison wasn't available, fall back to size comparison
        if not hash_comparison_available:
            # For files, try to get size if not already available
            if not hasattr(existing_file, 'size'):
                try:
                    print(f"[?] Fetching file size for comparison: {sanitized_name}")
                    # Try to refresh the item's properties
                    # Just use the existing_file object directly since we already have it
                    existing_file = existing_file.get().select(["size", "name", "file", "folder"]).execute_query()
                except Exception as select_error:
                    error_str = str(select_error)
                    # Check if this is the specific dangerous path error
                    if "dangerous Request.Path" in error_str or "%3Ckey%3E" in error_str:
                        print(f"[!] SharePoint API error, will re-upload to be safe: URL encoding issue")
                    else:
                        print(f"[!] Failed to get file metadata, will re-upload: {select_error}")
                    return True, False, None, local_hash
            # File exists - compare metadata
            # Try multiple ways to get size (different APIs use different property names)
            remote_size = None

            # Debug: Log what properties are available (verbose mode)
            if debug_metadata:
                print(f"[DEBUG] Available properties for {sanitized_name}:")
                print(f"  - Has 'size' attr: {hasattr(existing_file, 'size')}, value: {getattr(existing_file, 'size', 'N/A')}")
                print(f"  - Has 'length' attr: {hasattr(existing_file, 'length')}, value: {getattr(existing_file, 'length', 'N/A')}")
                print(f"  - Has 'properties' dict: {hasattr(existing_file, 'properties')}")

                if hasattr(existing_file, 'properties') and existing_file.properties:
                    print(f"  - Properties dict keys: {list(existing_file.properties.keys())[:10]}...")  # First 10 keys

            # Try Graph API DriveItem properties
            if hasattr(existing_file, 'size') and existing_file.size is not None:
                remote_size = existing_file.size
                if debug_metadata:
                    print(f"[DEBUG] Got size from 'size' property: {remote_size}")
            # Try SharePoint File properties
            elif hasattr(existing_file, 'length') and existing_file.length is not None:
                remote_size = existing_file.length
                if debug_metadata:
                    print(f"[DEBUG] Got size from 'length' property: {remote_size}")
            # Try properties dictionary (dynamic properties)
            elif hasattr(existing_file, 'properties'):
                remote_size = existing_file.properties.get('size') or existing_file.properties.get('Size') or existing_file.properties.get('length') or existing_file.properties.get('Length')
                if remote_size and debug_metadata:
                    print(f"[DEBUG] Got size from properties dict: {remote_size}")

            if remote_size is None:
                # If we still can't get size, log detailed info
                print(f"[!] Cannot determine remote file size for: {sanitized_name}")
                print(f"[DEBUG] Object type: {type(existing_file).__name__}")
                print(f"[DEBUG] Object attributes: {[attr for attr in dir(existing_file) if not attr.startswith('_')][:20]}...")
                return True, True, existing_file, local_hash

            # Compare file sizes only (hash comparison not available)
            size_matches = (local_size == remote_size)
            needs_update = not size_matches

            if not needs_update:
                print(f"[=] File unchanged (size: {local_size:,} bytes): {sanitized_name}")
                if upload_stats_dict:
                    upload_stats_dict['skipped_files'] += 1
                    upload_stats_dict['bytes_skipped'] += local_size
            else:
                print(f"[*] File size changed (local: {local_size:,} vs remote: {remote_size:,}): {sanitized_name}")

            return needs_update, True, existing_file, local_hash
        else:
            # Item exists but it's not a file or folder we can identify
            print(f"[?] Unable to determine type of existing item: {sanitized_name}")
            return True, False, None, local_hash

    except Exception as e:
        # File doesn't exist in SharePoint (404 error is expected)
        # Check if it's actually a 404 or another error
        error_str = str(e)
        if "404" in error_str or "not found" in error_str.lower() or "itemNotFound" in error_str:
            print(f"[+] New file to upload: {sanitized_name}")
        else:
            # Some other error occurred
            print(f"[?] Error checking file existence: {e}")
            print(f"[DEBUG] Error type: {type(e).__name__}")
            print(f"[DEBUG] Full error: {error_str[:500]}")  # First 500 chars
            print(f"[+] Assuming new file: {sanitized_name}")
        return True, False, None, local_hash
