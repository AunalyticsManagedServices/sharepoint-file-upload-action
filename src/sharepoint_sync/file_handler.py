# -*- coding: utf-8 -*-
"""
File handling operations for SharePoint sync.

This module provides functions for file sanitization, hashing, comparison, and exclusion.
"""

import os
import xxhash
import fnmatch
from .utils import is_debug_enabled


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
        if is_debug_enabled():
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
        if is_debug_enabled():
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


def check_file_needs_update(local_path, file_name, site_url, list_name, filehash_column_available,
                            tenant_id=None, client_id=None, client_secret=None, login_endpoint=None,
                            graph_endpoint=None, upload_stats_dict=None, pre_calculated_hash=None, display_path=None):
    """
    Check if a file in SharePoint needs to be updated by comparing hash or size.

    This function implements efficient file comparison to avoid unnecessary uploads.
    Files are compared using:
    1. FileHash (xxHash128) if column exists - most reliable
    2. Size comparison as fallback - works without custom columns

    Args:
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
        pre_calculated_hash (str, optional): Pre-calculated hash to use instead of calculating from file
                                             (useful for converted markdown where source .md hash is used)
        display_path (str, optional): Relative path for display in debug output (e.g., 'docs/api/README.html')
                                     If not provided, falls back to sanitized_name

    Returns:
        tuple: (needs_update: bool, exists: bool, remote_file: None, local_hash: str or None)
            - needs_update: True if file should be uploaded
            - exists: True if file exists in SharePoint
            - remote_file: Always None (no longer using Office365 DriveItem objects)
            - local_hash: The calculated or provided hash of the file

    Example:
        needs_update, exists, remote, hash_val = check_file_needs_update(
            "/path/to/file.pdf", "file.pdf", "site.sharepoint.com", "Documents", True,
            tenant_id, client_id, client_secret, login_endpoint, graph_endpoint
        )
        if not needs_update:
            print("File is up to date, skipping")
    """

    # Sanitize the file name to match what would be stored in SharePoint
    sanitized_name = sanitize_sharepoint_name(file_name, is_folder=False)

    # Use pre-calculated hash if provided, otherwise calculate from file
    local_hash = None
    if pre_calculated_hash:
        local_hash = pre_calculated_hash
        if is_debug_enabled():
            print(f"[#] Using pre-calculated hash: {local_hash[:8]}... for {sanitized_name}")
    else:
        local_hash = calculate_file_hash(local_path)
        if local_hash:
            if is_debug_enabled():
                print(f"[#] Local hash: {local_hash[:8]}... for {sanitized_name}")

    # Get local file information
    local_size = os.path.getsize(local_path)

    # Get debug flag (used throughout function)
    debug_metadata = os.environ.get('DEBUG_METADATA', 'false').lower() == 'true'

    # Debug: Show what we're checking
    if is_debug_enabled():
        display_name = display_path if display_path else sanitized_name
        print(f"[?] Checking if file exists in SharePoint: {display_name}")

    # Use Graph REST API to check file existence and get metadata
    # This replaces the Office365 library usage
    try:
        # Try to get the FileHash property and other metadata using direct REST API
        hash_comparison_available = False
        file_exists = False
        remote_size = None

        # Try to get file metadata using Graph REST API
        if all([tenant_id, client_id, client_secret, login_endpoint, graph_endpoint]):
            try:
                # Use direct Graph API REST calls to get SharePoint list item with custom columns
                from .graph_api import get_sharepoint_list_item_by_filename

                list_item_data = get_sharepoint_list_item_by_filename(
                    site_url, list_name, sanitized_name,
                    tenant_id, client_id, client_secret, login_endpoint, graph_endpoint
                )

                if list_item_data and 'fields' in list_item_data:
                    file_exists = True  # File found in SharePoint
                    fields = list_item_data['fields']

                    if debug_metadata:
                        print(f"[DEBUG] Retrieving metadata for {sanitized_name}")
                        print(f"[DEBUG] Available field properties: {list(fields.keys())}")

                    # Get file size if available
                    remote_size = fields.get('FileSizeDisplay') or fields.get('File_x0020_Size')
                    if isinstance(remote_size, str):
                        try:
                            remote_size = int(remote_size)
                        except (ValueError, TypeError):
                            remote_size = None

                    # Try to get FileHash if column is available
                    if filehash_column_available:
                        remote_hash = fields.get('FileHash')

                        if remote_hash:
                            hash_comparison_available = True
                            if is_debug_enabled():
                                print(f"[#] Remote hash: {remote_hash[:8]}... for {sanitized_name}")

                            # Compare hashes - this is the most reliable comparison
                            if upload_stats_dict:
                                upload_stats_dict['compared_by_hash'] = upload_stats_dict.get('compared_by_hash', 0) + 1

                            if local_hash and local_hash == remote_hash:
                                if is_debug_enabled():
                                    print(f"[=] File unchanged (hash match): {sanitized_name}")
                                if upload_stats_dict:
                                    upload_stats_dict['skipped_files'] += 1
                                    upload_stats_dict['bytes_skipped'] += local_size
                                    upload_stats_dict['hash_matched'] = upload_stats_dict.get('hash_matched', 0) + 1
                                return False, True, None, local_hash
                            elif local_hash:
                                if is_debug_enabled():
                                    print(f"[*] File changed (hash mismatch): {sanitized_name}")
                                return True, True, None, local_hash
                        elif debug_metadata:
                            print(f"[DEBUG] FileHash not found in list item fields")
                elif debug_metadata:
                    print(f"[DEBUG] Could not retrieve list item data for {sanitized_name}")

            except Exception as api_error:
                # File might not exist, or we can't access it
                if is_debug_enabled():
                    print(f"[!] Could not retrieve file metadata via REST API: {str(api_error)[:100]}")
                file_exists = False
                hash_comparison_available = False

        # If file doesn't exist, needs upload
        if not file_exists:
            if is_debug_enabled():
                print(f"[+] New file to upload: {sanitized_name}")
            return True, False, None, local_hash

        # If hash comparison wasn't available, fall back to size comparison
        if not hash_comparison_available:
            if debug_metadata:
                print(f"[DEBUG] FileHash not available, using size comparison")

            if remote_size is None:
                # If we still can't get size, assume file needs update
                if is_debug_enabled():
                    print(f"[!] Cannot determine remote file size for: {sanitized_name}")
                return True, True, None, local_hash

            # Compare file sizes only (hash comparison not available)
            if upload_stats_dict:
                upload_stats_dict['compared_by_size'] = upload_stats_dict.get('compared_by_size', 0) + 1

            size_matches = (local_size == remote_size)
            needs_update = not size_matches

            if not needs_update:
                if is_debug_enabled():
                    print(f"[=] File unchanged (size: {local_size:,} bytes): {sanitized_name}")
                if upload_stats_dict:
                    upload_stats_dict['skipped_files'] += 1
                    upload_stats_dict['bytes_skipped'] += local_size
            else:
                if is_debug_enabled():
                    display_name = display_path if display_path else sanitized_name
                    print(f"[*] File size changed (local: {local_size:,} vs remote: {remote_size:,}): {display_name}")

            return needs_update, True, None, local_hash

        # Should not reach here, but return safe default
        return True, file_exists, None, local_hash

    except Exception as e:
        # File doesn't exist in SharePoint (404 error is expected)
        # Check if it's actually a 404 or another error
        error_str = str(e)
        if "404" in error_str or "not found" in error_str.lower() or "itemNotFound" in error_str:
            if is_debug_enabled():
                print(f"[+] New file to upload: {sanitized_name}")
        else:
            # Some other error occurred
            print(f"[?] Error checking file existence: {e}")
            print(f"[DEBUG] Error type: {type(e).__name__}")
            print(f"[DEBUG] Full error: {error_str[:500]}")  # First 500 chars
            print(f"[+] Assuming new file: {sanitized_name}")
        return True, False, None, local_hash


def check_files_need_update_parallel(file_list, site_url, list_name,
                                     filehash_available, tenant_id, client_id,
                                     client_secret, login_endpoint, graph_endpoint,
                                     upload_stats_dict, max_workers=10):
    """
    Check multiple files concurrently to determine which need uploading.

    Performs parallel existence/change checks to build upload queue faster.
    Particularly useful when processing large numbers of files.

    Args:
        file_list (list): List of file paths to check
        site_url (str): SharePoint site URL
        list_name (str): SharePoint library name
        filehash_available (bool): Whether FileHash column exists
        tenant_id (str): Azure AD tenant ID
        client_id (str): Azure AD client ID
        client_secret (str): Azure AD client secret
        login_endpoint (str): Azure AD endpoint
        graph_endpoint (str): Graph API endpoint
        upload_stats_dict (dict): Upload statistics dictionary
        max_workers (int): Maximum concurrent checks (default: 10)

    Returns:
        dict: Mapping of {file_path: (needs_update, exists, remote_file, local_hash)}

    Example:
        check_results = check_files_need_update_parallel(
        ...     files, site_url, lib_name, True, ...
        ... )
        files_to_upload = [f for f, (needs_update, _, _, _) in check_results.items() if needs_update]

    Note:
        - 2-4x faster than sequential checks
        - Thread-safe statistics updates via locking
        - Useful for force_upload=False mode
    """
    from concurrent.futures import ThreadPoolExecutor, as_completed
    from .thread_utils import ThreadSafeStatsWrapper
    import threading

    results = {}
    results_lock = threading.Lock()

    # Wrap stats for thread safety
    stats_wrapper = ThreadSafeStatsWrapper(upload_stats_dict)

    def check_single_file(file_path):
        """Worker function to check single file"""
        file_name = os.path.basename(file_path)

        result = check_file_needs_update(
            file_path, file_name, site_url, list_name,
            filehash_available, tenant_id, client_id, client_secret,
            login_endpoint, graph_endpoint, stats_wrapper
        )

        with results_lock:
            results[file_path] = result

    # Execute checks in parallel
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(check_single_file, f) for f in file_list]

        # Wait for all to complete
        for future in as_completed(futures):
            try:
                future.result()
            except Exception as e:
                # Errors already logged by check_file_needs_update
                if is_debug_enabled():
                    print(f"[!] File check error: {e}")

    return results
