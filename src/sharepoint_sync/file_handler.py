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

    except FileNotFoundError:
        # File was deleted or moved during sync
        print(f"[!] ========================================")
        print(f"[!] FILE NOT FOUND")
        print(f"[!] ========================================")
        print(f"[!] File: {file_path}")
        print(f"[!] ")
        print(f"[!] File may have been deleted or moved during sync operation.")
        print(f"[!] ")
        print(f"[!] Troubleshooting:")
        print(f"[!]   - Verify file exists before running sync")
        print(f"[!]   - Check if file was moved by another process")
        print(f"[!]   - Exclude this file if it's temporary or auto-generated")
        print(f"[!] ========================================")
        return None

    except PermissionError:
        # Cannot read file due to permissions
        print(f"[!] ========================================")
        print(f"[!] PERMISSION DENIED")
        print(f"[!] ========================================")
        print(f"[!] File: {file_path}")
        print(f"[!] ")
        print(f"[!] Cannot read file - permission denied.")
        print(f"[!] ")
        print(f"[!] Troubleshooting:")
        print(f"[!]   1. Verify file permissions allow reading")
        print(f"[!]   2. Check if file is locked by another process")
        print(f"[!]   3. On Windows, check if file is opened exclusively by another app")
        print(f"[!]   4. Run with appropriate permissions if needed")
        print(f"[!]   5. Consider excluding this file from sync")
        print(f"[!] ========================================")
        return None

    except OSError as e:
        # I/O errors (disk issues, network drive problems, etc.)
        print(f"[!] ========================================")
        print(f"[!] FILE I/O ERROR")
        print(f"[!] ========================================")
        print(f"[!] File: {file_path}")
        print(f"[!] ")
        print(f"[!] Could not read file due to I/O error.")
        print(f"[!] ")
        print(f"[!] Troubleshooting:")
        print(f"[!]   1. Check disk health if errors persist (run: chkdsk on Windows, fsck on Linux)")
        print(f"[!]   2. Verify network drive connectivity if file is on network share")
        print(f"[!]   3. Check available disk space (may be full)")
        print(f"[!]   4. Verify filesystem is not corrupted")
        print(f"[!] ")
        print(f"[!] Technical details: {str(e)[:200]}")
        print(f"[!] ========================================")
        return None

    except MemoryError:
        # Out of memory - file may be extremely large
        file_size_mb = os.path.getsize(file_path) / (1024 * 1024) if os.path.exists(file_path) else 0
        print(f"[!] ========================================")
        print(f"[!] OUT OF MEMORY")
        print(f"[!] ========================================")
        print(f"[!] File: {file_path}")
        print(f"[!] File size: {file_size_mb:.2f} MB")
        print(f"[!] ")
        print(f"[!] Ran out of memory while calculating hash.")
        print(f"[!] ")
        print(f"[!] Troubleshooting:")
        print(f"[!]   1. File may be extremely large")
        print(f"[!]   2. Increase available memory for Docker container")
        print(f"[!]   3. Close other memory-intensive processes")
        print(f"[!]   4. Consider excluding very large files from sync")
        print(f"[!] ")
        print(f"[!] Note: Hash calculation uses dynamic chunk sizing (64KB-8MB)")
        print(f"[!]       to minimize memory usage, but very large files may still")
        print(f"[!]       cause issues on low-memory systems.")
        print(f"[!] ========================================")
        return None

    except UnicodeDecodeError as e:
        # File path has encoding issues (rare but possible)
        if is_debug_enabled():
            print(f"[!] ========================================")
            print(f"[!] FILE PATH ENCODING ERROR")
            print(f"[!] ========================================")
            print(f"[!] File path contains characters that cannot be decoded.")
            print(f"[!] ")
            print(f"[!] Troubleshooting:")
            print(f"[!]   1. File path may contain non-UTF-8 characters")
            print(f"[!]   2. Rename file to use standard ASCII characters")
            print(f"[!]   3. Check filesystem encoding settings")
            print(f"[!] ")
            print(f"[!] Technical details: {str(e)[:200]}")
            print(f"[!] ========================================")
        return None

    except Exception as e:
        # Unexpected errors - show detailed info in debug mode
        if is_debug_enabled():
            print(f"[!] Unexpected error calculating hash for {file_path}")
            print(f"    Error type: {type(e).__name__}")
            print(f"    Error: {str(e)[:200]}")
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
                            graph_endpoint=None, upload_stats_dict=None, pre_calculated_hash=None, display_path=None,
                            site_id=None, drive_id=None, parent_item_id=None, sharepoint_cache=None):
    """
    Check if a file in SharePoint needs to be updated by comparing hash or size.

    This function implements efficient file comparison to avoid unnecessary uploads.
    Files are compared using:
    1. Cache lookup (if cache provided) - fastest, no API calls
    2. FileHash (xxHash128) via API if column exists - most reliable
    3. Size comparison as fallback - works without custom columns

    Performance:
        - With cache: Instant lookup, 0 API calls
        - Without cache: 1 API call per file check

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
        site_id (str, optional): SharePoint site ID for path-based queries (preferred method)
        drive_id (str, optional): SharePoint drive ID for path-based queries (preferred method)
        parent_item_id (str, optional): Parent folder item ID for path-based queries (preferred method)
        sharepoint_cache (dict, optional): Pre-built cache of SharePoint file metadata
                                          Format: {"path/to/file.html": {"file_hash": "...", "size": 123, ...}}
                                          If None, falls back to individual API queries

    Returns:
        tuple: (needs_update: bool, exists: bool, remote_file: None, local_hash: str or None)
            - needs_update: True if file should be uploaded
            - exists: True if file exists in SharePoint
            - remote_file: Always None (no longer using Office365 DriveItem objects)
            - local_hash: The calculated or provided hash of the file

    Example:
        # With cache (recommended for bulk operations)
        cache = build_sharepoint_cache(...)
        needs_update, exists, remote, hash_val = check_file_needs_update(
            "/path/to/file.pdf", "file.pdf", "site.sharepoint.com", "Documents", True,
            sharepoint_cache=cache
        )

        # Without cache (falls back to API)
        needs_update, exists, remote, hash_val = check_file_needs_update(
            "/path/to/file.pdf", "file.pdf", "site.sharepoint.com", "Documents", True,
            tenant_id, client_id, client_secret, login_endpoint, graph_endpoint,
            site_id=site_id, drive_id=drive_id, parent_item_id=parent_item_id
        )
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

    # ============================================================================
    # CACHE LOOKUP (if available) - fastest path, no API calls
    # ============================================================================
    if sharepoint_cache is not None and display_path:
        # Try cache lookup using display_path (relative path)
        cached_file = sharepoint_cache.get(display_path)

        if cached_file:
            # Cache hit! Use cached metadata instead of API call
            if upload_stats_dict:
                if hasattr(upload_stats_dict, 'increment'):
                    upload_stats_dict.increment('cache_hits')
                else:
                    upload_stats_dict['cache_hits'] = upload_stats_dict.get('cache_hits', 0) + 1

            if is_debug_enabled():
                print(f"[CACHE HIT] Found {display_path} in cache")

            cached_hash = cached_file.get('file_hash')
            cached_size = cached_file.get('size')
            list_item_id = cached_file.get('list_item_id')

            # Try hash comparison first if available
            if filehash_column_available and cached_hash and local_hash:
                if upload_stats_dict:
                    if hasattr(upload_stats_dict, 'increment'):
                        upload_stats_dict.increment('compared_by_hash')
                    else:
                        upload_stats_dict['compared_by_hash'] = upload_stats_dict.get('compared_by_hash', 0) + 1

                if cached_hash == local_hash:
                    # Hash match - file unchanged
                    if is_debug_enabled():
                        print(f"[=] File unchanged (cached hash match): {display_path}")
                    if upload_stats_dict:
                        upload_stats_dict['skipped_files'] += 1
                        upload_stats_dict['bytes_skipped'] += local_size
                        if hasattr(upload_stats_dict, 'increment'):
                            upload_stats_dict.increment('hash_matched')
                        else:
                            upload_stats_dict['hash_matched'] = upload_stats_dict.get('hash_matched', 0) + 1
                    return False, True, None, local_hash
                else:
                    # Hash mismatch - file changed
                    if is_debug_enabled():
                        print(f"[*] File changed (cached hash mismatch): {display_path}")
                    return True, True, None, local_hash

            # Fall back to size comparison if hash not available
            elif cached_size is not None:
                if upload_stats_dict:
                    if hasattr(upload_stats_dict, 'increment'):
                        upload_stats_dict.increment('compared_by_size')
                    else:
                        upload_stats_dict['compared_by_size'] = upload_stats_dict.get('compared_by_size', 0) + 1

                if cached_size == local_size:
                    # Size match - likely unchanged
                    if is_debug_enabled():
                        print(f"[=] File unchanged (cached size match): {display_path}")
                    if upload_stats_dict:
                        upload_stats_dict['skipped_files'] += 1
                        upload_stats_dict['bytes_skipped'] += local_size

                    # Backfill empty FileHash if column exists
                    if (filehash_column_available and not cached_hash and local_hash and
                        list_item_id and site_url and list_name):
                        if is_debug_enabled():
                            print(f"[#] Backfilling empty FileHash for cached file: {display_path}")
                        try:
                            from .graph_api import update_sharepoint_list_item_field
                            success = update_sharepoint_list_item_field(
                                site_url, list_name, list_item_id, 'FileHash', local_hash,
                                tenant_id, client_id, client_secret, login_endpoint, graph_endpoint
                            )
                            if success:
                                if is_debug_enabled():
                                    print(f"[✓] FileHash backfilled: {local_hash[:8]}...")
                                if upload_stats_dict:
                                    if hasattr(upload_stats_dict, 'increment'):
                                        upload_stats_dict.increment('hash_backfilled')
                                    else:
                                        upload_stats_dict['hash_backfilled'] = upload_stats_dict.get('hash_backfilled', 0) + 1
                            else:
                                if upload_stats_dict:
                                    if hasattr(upload_stats_dict, 'increment'):
                                        upload_stats_dict.increment('hash_backfill_failed')
                                    else:
                                        upload_stats_dict['hash_backfill_failed'] = upload_stats_dict.get('hash_backfill_failed', 0) + 1
                        except Exception:
                            if upload_stats_dict:
                                if hasattr(upload_stats_dict, 'increment'):
                                    upload_stats_dict.increment('hash_backfill_failed')
                                else:
                                    upload_stats_dict['hash_backfill_failed'] = upload_stats_dict.get('hash_backfill_failed', 0) + 1

                    return False, True, None, local_hash
                else:
                    # Size mismatch - file changed
                    if is_debug_enabled():
                        print(f"[*] File changed (cached size mismatch): {display_path}")
                    return True, True, None, local_hash
        else:
            # Cache miss - file not found in cache
            # Fall through to API query to verify file status (safer than assuming new)
            if upload_stats_dict:
                if hasattr(upload_stats_dict, 'increment'):
                    upload_stats_dict.increment('cache_misses')
                else:
                    upload_stats_dict['cache_misses'] = upload_stats_dict.get('cache_misses', 0) + 1

            if is_debug_enabled():
                print(f"[CACHE MISS] {display_path} not found in cache - verifying with API query")
            # Don't return - fall through to API query below for safety

    # ============================================================================
    # FALLBACK: Individual API query (cache miss or cache not available)
    # ============================================================================
    # Track API query (fallback when cache not available)
    if upload_stats_dict:
        if hasattr(upload_stats_dict, 'increment'):
            upload_stats_dict.increment('api_queries')
        else:
            upload_stats_dict['api_queries'] = upload_stats_dict.get('api_queries', 0) + 1

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
                # Prefer path-based query (most reliable, especially for duplicate filenames)
                list_item_data = None
                if all([site_id, drive_id, parent_item_id]):
                    if is_debug_enabled():
                        print(f"[DEBUG] Querying by path: parent={parent_item_id}, file={sanitized_name}")

                    # Use path-based query to get exact file (fixes duplicate filename bug)
                    from .graph_api import get_drive_item_by_path_with_list_item

                    item_with_list = get_drive_item_by_path_with_list_item(
                        site_id, drive_id, parent_item_id, sanitized_name,
                        tenant_id, client_id, client_secret, login_endpoint, graph_endpoint
                    )

                    # Extract listItem fields from the response
                    if item_with_list and 'listItem' in item_with_list:
                        list_item_data = {
                            'fields': item_with_list['listItem'].get('fields', {})
                        }
                        if is_debug_enabled():
                            print(f"[DEBUG] Retrieved file metadata by path")

                # If path-based query failed, we cannot reliably check the file
                # (filename-only search is unreliable for duplicate names)
                if not list_item_data:
                    if is_debug_enabled():
                        print(f"[DEBUG] Could not retrieve file metadata by path")
                        print(f"[DEBUG] Missing required parameters: site_id={site_id is not None}, drive_id={drive_id is not None}, parent_item_id={parent_item_id is not None}")

                    # Without path-based query, we must assume file needs update
                    # This is safer than using unreliable filename-only search
                    if is_debug_enabled():
                        print(f"[!] Cannot verify file status, assuming needs update: {sanitized_name}")
                    return True, False, None, local_hash

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
                                # Use atomic increment if available (parallel mode), otherwise use get/set pattern
                                if hasattr(upload_stats_dict, 'increment'):
                                    upload_stats_dict.increment('compared_by_hash')
                                else:
                                    upload_stats_dict['compared_by_hash'] = upload_stats_dict.get('compared_by_hash', 0) + 1

                            if local_hash and local_hash == remote_hash:
                                if is_debug_enabled():
                                    print(f"[=] File unchanged (hash match): {sanitized_name}")
                                if upload_stats_dict:
                                    upload_stats_dict['skipped_files'] += 1
                                    upload_stats_dict['bytes_skipped'] += local_size
                                    # Use atomic increment if available (parallel mode), otherwise use get/set pattern
                                    if hasattr(upload_stats_dict, 'increment'):
                                        upload_stats_dict.increment('hash_matched')
                                    else:
                                        upload_stats_dict['hash_matched'] = upload_stats_dict.get('hash_matched', 0) + 1
                                return False, True, None, local_hash
                            elif local_hash:
                                if is_debug_enabled():
                                    print(f"[*] File changed (hash mismatch): {sanitized_name}")
                                return True, True, None, local_hash
                        else:
                            # FileHash column exists but value is empty for this file
                            if debug_metadata:
                                print(f"[DEBUG] FileHash not found in list item fields")
                            if upload_stats_dict:
                                # Use atomic increment if available (parallel mode), otherwise use get/set pattern
                                if hasattr(upload_stats_dict, 'increment'):
                                    upload_stats_dict.increment('hash_empty_found')
                                else:
                                    upload_stats_dict['hash_empty_found'] = upload_stats_dict.get('hash_empty_found', 0) + 1
                    else:
                        # FileHash column doesn't exist at all
                        if upload_stats_dict:
                            # Use atomic increment if available (parallel mode), otherwise use get/set pattern
                            if hasattr(upload_stats_dict, 'increment'):
                                upload_stats_dict.increment('hash_column_unavailable')
                            else:
                                upload_stats_dict['hash_column_unavailable'] = upload_stats_dict.get('hash_column_unavailable', 0) + 1
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
                # Use atomic increment if available (parallel mode), otherwise use get/set pattern
                if hasattr(upload_stats_dict, 'increment'):
                    upload_stats_dict.increment('compared_by_size')
                else:
                    upload_stats_dict['compared_by_size'] = upload_stats_dict.get('compared_by_size', 0) + 1

            size_matches = (local_size == remote_size)
            needs_update = not size_matches

            if not needs_update:
                if is_debug_enabled():
                    print(f"[=] File unchanged (size: {local_size:,} bytes): {sanitized_name}")
                if upload_stats_dict:
                    upload_stats_dict['skipped_files'] += 1
                    upload_stats_dict['bytes_skipped'] += local_size

                # Backfill empty FileHash values
                # If FileHash column exists but value is empty, and we have confirmed
                # file is unchanged via size comparison, update the hash without re-uploading
                if (filehash_column_available and
                    not hash_comparison_available and  # Hash was empty
                    local_hash and  # We have a calculated hash
                    site_url and list_name and  # Required for update
                    item_with_list and 'listItem' in item_with_list and 'id' in item_with_list['listItem']):

                    # Attempt to backfill the FileHash
                    item_id = item_with_list['listItem']['id']

                    if is_debug_enabled():
                        display_name = display_path if display_path else sanitized_name
                        print(f"[#] Backfilling empty FileHash for unchanged file: {display_name}")

                    try:
                        from .graph_api import update_sharepoint_list_item_field

                        success = update_sharepoint_list_item_field(
                            site_url, list_name, item_id, 'FileHash', local_hash,
                            tenant_id, client_id, client_secret, login_endpoint, graph_endpoint
                        )

                        if success:
                            if is_debug_enabled():
                                print(f"[✓] FileHash backfilled: {local_hash[:8]}...")
                            if upload_stats_dict:
                                # Use atomic increment if available (parallel mode), otherwise use get/set pattern
                                if hasattr(upload_stats_dict, 'increment'):
                                    upload_stats_dict.increment('hash_backfilled')
                                else:
                                    upload_stats_dict['hash_backfilled'] = upload_stats_dict.get('hash_backfilled', 0) + 1
                        else:
                            if is_debug_enabled():
                                print(f"[!] Failed to backfill FileHash")
                            if upload_stats_dict:
                                # Use atomic increment if available (parallel mode), otherwise use get/set pattern
                                if hasattr(upload_stats_dict, 'increment'):
                                    upload_stats_dict.increment('hash_backfill_failed')
                                else:
                                    upload_stats_dict['hash_backfill_failed'] = upload_stats_dict.get('hash_backfill_failed', 0) + 1

                    except Exception as backfill_error:
                        if is_debug_enabled():
                            print(f"[!] Error backfilling FileHash: {str(backfill_error)[:200]}")
                        if upload_stats_dict:
                            # Use atomic increment if available (parallel mode), otherwise use get/set pattern
                            if hasattr(upload_stats_dict, 'increment'):
                                upload_stats_dict.increment('hash_backfill_failed')
                            else:
                                upload_stats_dict['hash_backfill_failed'] = upload_stats_dict.get('hash_backfill_failed', 0) + 1

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
