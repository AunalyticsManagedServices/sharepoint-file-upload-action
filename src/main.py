#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SharePoint File Upload Script for GitHub Actions
================================================
Compatible with Office365-REST-Python-Client v2.6.2+

PURPOSE:
    This script automates file uploads from GitHub repositories to SharePoint/OneDrive,
    typically used in CI/CD pipelines to sync documentation, reports, or build artifacts.

SYNOPSIS:
    python main.py <site_name> <sharepoint_host> <tenant_id>
                                 <client_id> <client_secret> <upload_path>
                                 <file_path> [max_retry] [login_endpoint]
                                 [graph_endpoint] [recursive] [force_upload]
                                 [convert_md_to_html] [exclude_patterns]

PARAMETERS:
    Required Parameters:
    -------------------
    <site_name>
        SharePoint site name from your site URL.
        `Example`: For 'https://company.sharepoint.com/sites/TeamSite', use 'TeamSite'
        `Type`: String
        `Position`: 1

    <sharepoint_host>
        SharePoint tenant domain name.
        `Example`: 'company.sharepoint.com' or 'company-my.sharepoint.com'
        For GovCloud: 'company.sharepoint.us'
        `Type`: String (FQDN)
        `Position`: 2

    <tenant_id>
        Azure AD tenant ID (GUID format).
        Find in Azure Portal → Azure Active Directory → Properties → Tenant ID
        `Example`: '12345678-1234-1234-1234-123456789abc'
        `Type`: String (GUID)
        `Position`: 3

    <client_id>
        Azure AD App Registration application (client) ID.
        Find in Azure Portal → App Registrations → Your App → Application ID
        Requires Sites.ReadWrite.All (or Sites.Manage.All for column creation)
        `Example`: '87654321-4321-4321-4321-cba987654321'
        `Type`: String (GUID)
        `Position`: 4

    <client_secret>
        Azure AD App Registration client secret value.
        Create in Azure Portal → App Registrations → Certificates & secrets
        `WARNING: Keep this secure! Never commit to version control.
        Store in GitHub Secrets or environment variables.
        `Type`: String (sensitive)
        `Position`: 5

    <upload_path>
        Target path in SharePoint document library where files will be uploaded.
        `Format`: 'LibraryName/Folder/Subfolder' (use forward slashes)
        `Example`: 'Documents/Reports/2024' or 'Shared Documents/Archive'
        Creates missing folders automatically.
        `Type`: String (path)
        `Position`: 6

    <file_path>
        Local file or glob pattern to upload.
        Supports wildcards: *, ?, [seq], [!seq]

        Path Separators:
            - Forward slashes (/) recommended for cross-platform compatibility
            - Backslashes (\\) work on Windows (use double backslash or raw strings)
            - Both absolute and relative paths supported

        Examples:
            Unix/Cross-platform style:
                - '*.pdf' (all PDFs in current directory)
                - 'docs/**/*.md' (all markdown files, requires recursive=True)
                - 'report.xlsx' (single file)
                - './build/artifacts/*' (all files in artifacts folder)

            Windows style:
                - '*.pdf' (all PDFs in current directory)
                - 'docs\\**\\*.md' (all markdown files recursively)
                - 'C:\\Reports\\*.xlsx' (absolute Windows path)
                - '.\\build\\artifacts\\*' (relative Windows path)

        `Type`: String (file path or glob pattern)
        `Position`: 7

    Optional Parameters:
    -------------------
    [max_retry]
        Maximum number of retry attempts for failed uploads.
        Default: 3
        Range: 0-10 (0 = no retries)
        Applies to network errors, timeouts, and transient server errors (5xx).
        `Type`: Integer
        `Position`: 8

    [login_endpoint]
        Azure AD authentication endpoint for special cloud environments.
        Default: 'login.microsoftonline.com' (Commercial Cloud)
        Other options:
            - 'login.microsoftonline.us' (US Government Cloud)
            - 'login.microsoftonline.de' (Germany Cloud)
            - 'login.chinacloudapi.cn' (China Cloud)
        `Type`: String (FQDN)
        `Position`: 9

    [graph_endpoint]
        Microsoft Graph API endpoint for special cloud environments.
        Default: 'graph.microsoft.com' (Commercial Cloud)
        Other options:
            - 'graph.microsoft.us' (US Government Cloud)
            - 'graph.microsoft.de' (Germany Cloud)
            - 'microsoftgraph.chinacloudapi.cn' (China Cloud)
        `Type`: String (FQDN)
        `Position`: 10

    [recursive]
        Enable recursive file matching for glob patterns with '**'.
        Default: 'False'
        Values: 'True' or 'False' (case-sensitive string)
        When True, patterns like 'docs/**/*.md' match files in all subdirectories.
        When False, only matches files in the specified directory.
        `Type`: String ('True'/'False')
        `Position`: 11

    [force_upload]
        Force upload all files, skipping hash/size comparison.
        Default: 'False'
        Values: 'True' or 'False' (case-sensitive string)
        When True, uploads all files regardless of changes (slower, more bandwidth).
        When False, uses smart sync with xxHash128 comparison (faster, efficient).
        Use cases: force refresh, corrupted files, testing.
        `Type`: String ('True'/'False')
        `Position`: 12

    [convert_md_to_html]
        Convert Markdown (.md) files to HTML with embedded Mermaid SVG diagrams.
        Default: 'True'
        Values: 'True' or 'False' (case-sensitive string)
        When True, converts .md → .html with GitHub-flavored styling and Mermaid rendering.
        When False, uploads .md files as-is (raw markdown).
        Requires: Node.js and @mermaid-js/mermaid-cli for diagram conversion.
        `Type`: String ('True'/'False')
        `Position`: 13

    [exclude_patterns]
        Comma-separated list of file/directory exclusion patterns.
        Default: '' (empty string - no exclusions)
        `Format`: 'pattern1,pattern2,pattern3' (comma-separated, no spaces around commas)

        Pattern Types:
            - Wildcard patterns: '*.tmp', '*.log', '*.pyc'
            - Directory names: '__pycache__', '.git', 'node_modules', '.svn'
            - Extension only: 'tmp', 'log' (automatically becomes '*.tmp', '*.log')
            - Specific files: 'config.local.json', 'secrets.txt'

        Pattern Matching:
            - Matches against filename/directory name (basename)
            - Matches against full path for precise exclusions
            - Directory names matched anywhere in path (e.g., '__pycache__' excludes all)
            - Case-sensitive on Linux, case-insensitive on Windows

        Cross-Platform Compatibility:
            - Works with both forward slashes (/) and backslashes (\\)
            - Path normalization ensures consistent matching on all platforms

        Common Use Cases:
            - Temporary files: '*.tmp,*.bak,*.swp'
            - Log files: '*.log'
            - Python artifacts: '*.pyc,__pycache__,.pytest_cache'
            - Node.js artifacts: 'node_modules,.npm'
            - Version control: '.git,.svn,.hg'
            - IDE files: '.vscode,.idea,*.code-workspace'
            - OS files: '.DS_Store,Thumbs.db,desktop.ini'

        Examples:
            - Exclude temp files: '*.tmp,*.bak'
            - Exclude Python cache: '__pycache__,*.pyc'
            - Exclude multiple types: '*.tmp,*.log,__pycache__,node_modules,.git'

        `Type`: String (comma-separated patterns)
        `Position`: 14

DESCRIPTION:
    - Intelligently syncs files to SharePoint, skipping unchanged files
    - Compares file size and modification time to detect changes
    - Uploads only new or modified files, saving time and bandwidth
    - Converts Markdown files to HTML with embedded SVG diagrams
    - Renders Mermaid diagrams as static SVG for SharePoint compatibility
    - Handles large files (>4MB) using resumable upload sessions
    - Creates missing folders automatically in SharePoint
    - Provides detailed statistics on uploads, updates, and skipped files
    - Supports recursive file matching with glob patterns
    - Optional force mode to upload all files regardless of changes

EXAMPLES:
    1. Upload a single file with defaults:
       python send_to_sharepoint.py TeamSite company.sharepoint.com \\
              tenant-guid-here client-guid-here client-secret-here \\
              "Documents/Archive" "report.pdf"

    2. Upload all PDFs with smart sync (skips unchanged):
       python send_to_sharepoint.py TeamSite company.sharepoint.com \\
              tenant-guid-here client-guid-here client-secret-here \\
              "Documents/Reports" "*.pdf"

    3. Upload markdown files recursively with conversion to HTML:
       python send_to_sharepoint.py TeamSite company.sharepoint.com \\
              tenant-guid-here client-guid-here client-secret-here \\
              "Shared Documents/Docs" "docs/**/*.md" 3 \\
              login.microsoftonline.com graph.microsoft.com True False True

    4. Force upload all files (skip hash comparison):
       python send_to_sharepoint.py TeamSite company.sharepoint.com \\
              tenant-guid-here client-guid-here client-secret-here \\
              "Documents/Backup" "*.xlsx" 5 \\
              login.microsoftonline.com graph.microsoft.com False True False

    5. US Government Cloud environment:
       python send_to_sharepoint.py TeamSite company.sharepoint.us \\
              tenant-guid-here client-guid-here client-secret-here \\
              "Documents/Reports" "*.pdf" 3 \\
              login.microsoftonline.us graph.microsoft.us

    6. Upload build artifacts with custom retry count:
       python send_to_sharepoint.py TeamSite company.sharepoint.com \\
              tenant-guid-here client-guid-here client-secret-here \\
              "Documents/Builds/v1.2.3" "./dist/*" 10

    7. Windows absolute path with backslashes:
       python send_to_sharepoint.py TeamSite company.sharepoint.com ^
              tenant-guid-here client-guid-here client-secret-here ^
              "Documents/Reports" "C:\\Users\\Documents\\Reports\\*.xlsx"

    8. Upload all files recursively excluding temporary files:
       python send_to_sharepoint.py TeamSite company.sharepoint.com \\
              tenant-guid-here client-guid-here client-secret-here \\
              "Documents/Project" "project/**/*" 3 \\
              login.microsoftonline.com graph.microsoft.com True False True \\
              "*.tmp,*.bak,*.log"

    9. Upload Python project excluding cache and compiled files:
       python send_to_sharepoint.py TeamSite company.sharepoint.com \\
              tenant-guid-here client-guid-here client-secret-here \\
              "Documents/Python" "src/**/*" 3 \\
              login.microsoftonline.com graph.microsoft.com True False False \\
              "__pycache__,*.pyc,.pytest_cache,.tox"

    10. Upload all files excluding version control and IDE directories:
        python send_to_sharepoint.py TeamSite company.sharepoint.com \\
               tenant-guid-here client-guid-here client-secret-here \\
               "Documents/Source" "./**/*" 3 \\
               login.microsoftonline.com graph.microsoft.com True False False \\
               ".git,.svn,.hg,.vscode,.idea,node_modules"

    11. Windows path with exclusions (all files except logs and temps):
        python send_to_sharepoint.py TeamSite company.sharepoint.com ^
               tenant-guid-here client-guid-here client-secret-here ^
               "Documents/Data" "C:\\Projects\\Data\\**\\*" 3 ^
               login.microsoftonline.com graph.microsoft.com True False False ^
               "*.log,*.tmp,*.bak,Thumbs.db,.DS_Store"

    Note: On Windows CMD, use ^ for line continuation instead of \\

REQUIREMENTS:
    - Python 3.6 or higher
    - office365-rest-python-client >= 2.6.2
    - msal (Microsoft Authentication Library)
    - Azure AD Enterprise Application with Graph API Sites.ReadWrite.All permission

AUTHOR:
    Mark Newton

VERSION:
    3.0.0 - Added smart sync, markdown to HTML conversion with Mermaid SVG support
"""

# ====================================
# IMPORTS - External libraries needed
# ====================================

import sys
import os
import glob
import tempfile
import time

# SharePoint sync modules
from sharepoint_sync.config import parse_config
from sharepoint_sync.graph_api import (
    create_graph_client, check_and_create_filehash_column, get_sharepoint_list_item_by_filename,
    list_files_in_folder_recursive, delete_file_from_sharepoint
)
from sharepoint_sync.file_handler import should_exclude_path, calculate_file_hash, sanitize_path_components
from sharepoint_sync.uploader import upload_file_with_structure, upload_file, ensure_folder_exists
from sharepoint_sync.markdown_converter import convert_markdown_to_html
from sharepoint_sync.monitoring import upload_stats, print_rate_limiting_summary
from sharepoint_sync.utils import is_debug_enabled
from sharepoint_sync.parallel_uploader import ParallelUploader


# ====================================================================
# FILE DISCOVERY - Finding files to upload
# ====================================================================

def discover_files(file_path, recursive, exclude_patterns_list):
    """
    Discover files based on glob patterns and exclusion filters.

    This function finds all files matching the given pattern(s) and applies
    exclusion filters to remove unwanted files/directories.

    Args:
        file_path (str): File path or glob pattern to match
        recursive (bool): Enable recursive matching for '**' patterns
        exclude_patterns_list (list): List of exclusion patterns

    Returns:
        tuple: (list of file paths, list of directory paths)

    Process:
        1. Use glob.glob() to find all items matching the pattern
        2. Apply exclusion filters to remove unwanted items
        3. Separate files from directories
        4. Return both lists

    Examples:
        >>> discover_files('*.pdf', False, [])
        (['report.pdf', 'invoice.pdf'], [])

        >>> discover_files('src/**/*.py', True, ['__pycache__', '*.pyc'])
        (['src/main.py', 'src/utils.py'], [])
    """
    # Use glob to find all files/directories matching the pattern
    # glob.glob() returns a list of paths matching a pathname pattern
    # Examples: '*.txt' finds all .txt files, '**/*.py' finds all .py files recursively
    local_items_unfiltered = glob.glob(file_path, recursive=recursive)

    # Apply exclusion filters if provided
    if exclude_patterns_list:
        local_items = [item for item in local_items_unfiltered if not should_exclude_path(item, exclude_patterns_list)]
        excluded_count = len(local_items_unfiltered) - len(local_items)
        if excluded_count > 0 and is_debug_enabled():
            print(f"[=] Excluded {excluded_count} item(s) matching exclusion patterns")
    else:
        local_items = local_items_unfiltered

    # Exit with error if no matches found
    if not local_items:
        if exclude_patterns_list and local_items_unfiltered:
            print(f"[Error] All files matched by pattern '{file_path}' were excluded by filters")
            print(f"[Error] {len(local_items_unfiltered)} file(s) found but all matched exclusion patterns: {', '.join(exclude_patterns_list)}")
        else:
            print(f"[Error] No files or directories matched pattern: {file_path}")
        sys.exit(1)  # Exit code 1 indicates error to calling process (e.g., GitHub Actions)

    # Separate files from directories
    local_files = []  # Will contain paths to actual files
    local_dirs = []   # Will contain paths to directories

    # Iterate through each matched item and categorize it
    for item in local_items:
        if os.path.isfile(item):  # Check if path points to a file
            local_files.append(item)  # Add to files list
        elif os.path.isdir(item):  # Check if path points to a directory
            local_dirs.append(item)   # Add to directories list

    return local_files, local_dirs


# ====================================================================
# BASE PATH CALCULATION - For maintaining folder structure
# ====================================================================

def calculate_base_path(local_files, local_dirs):
    """
    Calculate the base path to strip from file paths when uploading.

    This preserves the relative folder structure when uploading to SharePoint.

    Args:
        local_files (list): List of file paths
        local_dirs (list): List of directory paths

    Returns:
        str: Base path to use for relative path calculation

    Examples:
        >>> calculate_base_path(['/a/b/file1.txt', '/a/b/c/file2.txt'], [])
        '/a/b'

        >>> calculate_base_path([], ['/home/user/docs'])
        '/home/user'
    """
    base_path = ""  # Initialize empty base path

    if local_dirs:
        # If directories were selected, use the parent of the first directory
        # Example: If uploading "/home/user/docs", base is "/home/user"
        base_path = os.path.dirname(local_dirs[0])
    elif local_files:
        # If only files were selected, find their common parent directory
        # os.path.commonpath() finds the longest common path prefix
        # Example: ["/a/b/file1.txt", "/a/b/c/file2.txt"] → "/a/b"
        base_path = os.path.dirname(os.path.commonpath(local_files))

    return base_path


# ====================================================================
# SUMMARY REPORT
# ====================================================================

def print_summary(total_files, whatif_mode=False):
    """
    Print final summary report with upload statistics.

    Args:
        total_files (int): Total number of files processed
        whatif_mode (bool): Whether sync deletion is in WhatIf mode
    """
    def format_bytes(bytes_val):
        """Convert bytes to human-readable format"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if bytes_val < 1024.0:
                return f"{bytes_val:.1f} {unit}"
            bytes_val /= 1024.0
        return f"{bytes_val:.1f} TB"

    # Create visual separator for better readability
    print("\n" + "="*60)
    print("[✓] SYNC PROCESS COMPLETED")
    print("="*60)

    # Show detailed statistics
    stats = upload_stats.stats
    print(f"[STATS] Sync Statistics:")
    print(f"   - New files uploaded:       {stats['new_files']:>6}")
    print(f"   - Files updated:            {stats['replaced_files']:>6}")
    print(f"   - Files skipped (unchanged):{stats['skipped_files']:>6}")

    # Show deleted files with WhatIf indicator if applicable
    if stats['deleted_files'] > 0:
        if whatif_mode:
            print(f"   - Files deleted (WhatIf):   {stats['deleted_files']:>6}")
        else:
            print(f"   - Files deleted:            {stats['deleted_files']:>6}")

    print(f"   - Failed uploads:           {stats['failed_files']:>6}")
    print(f"   - Total files processed:    {total_files:>6}")

    # Show comparison methods if files were compared
    total_compared = stats.get('compared_by_hash', 0) + stats.get('compared_by_size', 0)
    if total_compared > 0:
        print(f"\n[COMPARE] File Comparison Methods:")
        hash_count = stats.get('compared_by_hash', 0)
        size_count = stats.get('compared_by_size', 0)
        hash_pct = (hash_count / total_compared * 100) if total_compared > 0 else 0
        size_pct = (size_count / total_compared * 100) if total_compared > 0 else 0
        print(f"   - Compared by hash:         {hash_count:>6} ({hash_pct:.1f}%)")
        print(f"   - Compared by size:         {size_count:>6} ({size_pct:.1f}%)")

    # Show FileHash column statistics if any hash operations occurred
    total_hash_ops = (stats.get('hash_new_saved', 0) + stats.get('hash_updated', 0) +
                     stats.get('hash_matched', 0) + stats.get('hash_save_failed', 0))
    if total_hash_ops > 0:
        print(f"\n[HASH] FileHash Column Statistics:")
        if stats.get('hash_new_saved', 0) > 0:
            print(f"   - New hashes saved:         {stats.get('hash_new_saved', 0):>6}")
        if stats.get('hash_updated', 0) > 0:
            print(f"   - Hashes updated:           {stats.get('hash_updated', 0):>6}")
        if stats.get('hash_matched', 0) > 0:
            print(f"   - Hash matches (skipped):   {stats.get('hash_matched', 0):>6}")
        if stats.get('hash_save_failed', 0) > 0:
            print(f"   - Hash save failures:       {stats.get('hash_save_failed', 0):>6}")

    print(f"\n[DATA] Transfer Summary:")
    print(f"   - Data uploaded:   {format_bytes(stats['bytes_uploaded'])}")
    print(f"   - Data skipped:    {format_bytes(stats['bytes_skipped'])}")
    print(f"   - Total savings:   {format_bytes(stats['bytes_skipped'])} ({stats['skipped_files']} files not re-uploaded)")

    # Calculate efficiency percentage
    total_processed = stats['new_files'] + stats['replaced_files'] + stats['skipped_files']
    if total_processed > 0:
        efficiency = (stats['skipped_files'] / total_processed) * 100
        print(f"\n[EFFICIENCY] {efficiency:.1f}% of files were already up-to-date")

    print("="*60)

    # Display rate limiting statistics
    print_rate_limiting_summary()


# ====================================================================
# LIBRARY NAME EXTRACTION
# ====================================================================

def get_library_name_from_path(upload_path):
    """
    Extract the document library name from the upload path.

    Args:
        upload_path (str): Upload path like 'Documents/Reports/2024'

    Returns:
        str: Library name (e.g., 'Documents')
    """
    library_name = "Documents"  # Default document library name
    if upload_path and "/" in upload_path:
        # If upload_path starts with a library name, use it
        path_parts = upload_path.split("/")
        if path_parts[0]:
            library_name = path_parts[0]
    return library_name


# ====================================================================
# SYNC DELETION - Remove orphaned files from SharePoint
# ====================================================================

def identify_files_to_delete(sharepoint_files, local_files_set, base_path):
    """
    Identify SharePoint files that should be deleted (not in local sync set).

    Args:
        sharepoint_files (list): List of file dicts from SharePoint (from list_files_in_folder_recursive)
        local_files_set (set): Set of relative file paths from local sync set
        base_path (str): Base path used for folder structure

    Returns:
        list: List of file dicts that should be deleted

    Note:
        Compares SharePoint files with local sync set to identify orphaned files
        that no longer exist locally and should be deleted from SharePoint.
    """
    files_to_delete = []
    debug_enabled = is_debug_enabled()

    if debug_enabled:
        print(f"\n[DEBUG] Comparing {len(sharepoint_files)} SharePoint files with {len(local_files_set)} local files...")
        print(f"[DEBUG] SharePoint files:")
        for sp_file in sharepoint_files:
            print(f"  [SP] {sp_file['path']}")
        print(f"\n[DEBUG] Local files set:")
        for local_path in sorted(local_files_set):
            print(f"  [LOCAL] {local_path}")
        print(f"\n[DEBUG] Starting comparison...")

    for sp_file in sharepoint_files:
        # The path in SharePoint (relative to upload folder)
        sp_path = sp_file['path']

        # Check if this file exists in our local set
        if sp_path not in local_files_set:
            files_to_delete.append(sp_file)

            if debug_enabled:
                print(f"  [×] ORPHANED: {sp_path} (not in local sync set)")
        elif debug_enabled:
            print(f"  [✓] MATCHED: {sp_path}")

    return files_to_delete


def perform_sync_deletion(root_drive, local_files, base_path, config):
    """
    Delete files from SharePoint that are not in the local sync set.

    Args:
        root_drive: SharePoint Drive object representing the upload folder
        local_files (list): List of local file paths being synced
        base_path (str): Base path for maintaining folder structure
        config: Configuration object

    Returns:
        int: Number of files successfully deleted

    Safety measures:
        - Only deletes files within the sync target folder
        - Requires explicit sync_delete flag to be enabled
        - Compares full relative paths to avoid accidental deletions
        - Provides detailed logging of deletions
    """
    debug_enabled = is_debug_enabled()

    # Step 1: List all files currently in SharePoint folder
    print("\n[*] Listing files in SharePoint target folder...")
    try:
        sharepoint_files = list_files_in_folder_recursive(root_drive, config.upload_path)
        print(f"[OK] Found {len(sharepoint_files)} files in SharePoint")
    except Exception as e:
        print(f"[!] Failed to list SharePoint files: {str(e)}")
        return 0

    # Step 2: Build set of local file relative paths
    # Need to calculate the relative paths the same way upload does
    local_files_set = set()

    if debug_enabled:
        print(f"\n[DEBUG] Building local file set (base_path: {base_path})...")

    for local_file in local_files:
        # Calculate relative path from base_path (preserve folder structure!)
        # This MUST match how upload_file_with_structure calculates paths
        if base_path:
            try:
                rel_path = os.path.relpath(local_file, base_path)
            except ValueError:
                # On Windows, relpath fails if paths are on different drives
                # Fall back to absolute path calculation
                rel_path = local_file
        else:
            # No base_path means upload to root - use full path
            rel_path = local_file

        # Normalize path separators to forward slashes (SharePoint style)
        rel_path = rel_path.replace(os.sep, '/')
        rel_path = rel_path.replace('\\', '/')  # Extra safety for Windows paths

        # Handle markdown to HTML conversion
        if local_file.lower().endswith('.md') and config.convert_md_to_html:
            # If converting .md to .html, the SharePoint file will be .html
            rel_path = rel_path[:-3] + '.html'

        # Sanitize path to match how uploader sanitizes
        from sharepoint_sync.file_handler import sanitize_path_components
        rel_path = sanitize_path_components(rel_path)

        local_files_set.add(rel_path)

        if debug_enabled:
            print(f"  [+] Local: {local_file} → {rel_path}")

    if debug_enabled:
        print(f"\n[DEBUG] Local files set contains {len(local_files_set)} items")
        print(f"[DEBUG] SharePoint returned {len(sharepoint_files)} items")

    # Step 3: Identify files to delete
    files_to_delete = identify_files_to_delete(sharepoint_files, local_files_set, base_path)

    if not files_to_delete:
        print("[OK] No orphaned files to delete from SharePoint")
        return 0

    # Step 4: Delete orphaned files (or show what would be deleted in WhatIf mode)
    if config.sync_delete_whatif:
        print(f"\n[!] Found {len(files_to_delete)} orphaned files (WhatIf mode - no actual deletions will occur)")
    else:
        print(f"\n[!] Found {len(files_to_delete)} orphaned files to delete from SharePoint")

    deleted_count = 0
    for file_info in files_to_delete:
        try:
            success = delete_file_from_sharepoint(
                file_info['drive_item'],
                file_info['path'],
                whatif=config.sync_delete_whatif
            )
            if success:
                deleted_count += 1
                upload_stats.stats['deleted_files'] += 1
        except Exception as e:
            print(f"[!] Error deleting {file_info['path']}: {str(e)}")

    if config.sync_delete_whatif:
        print(f"[✓] WhatIf: Would delete {deleted_count} orphaned files from SharePoint")
    else:
        print(f"[✓] Successfully deleted {deleted_count} orphaned files from SharePoint")
    return deleted_count


# ====================================================================
# MAIN EXECUTION
# ====================================================================

def main():
    """
    Main execution function that orchestrates the SharePoint sync process.

    Process:
        1. Parse configuration from command-line arguments
        2. Discover files matching the pattern
        3. Connect to SharePoint and verify FileHash column
        4. Calculate base path for folder structure preservation
        5. Process each file (with markdown conversion if enabled)
        6. Print summary statistics
        7. Exit with appropriate code
    """
    # Parse configuration from command-line arguments
    config = parse_config()

    # Display system configuration stats box
    print("\n" + "="*60)
    print("[✓] SYSTEM CONFIGURATION")
    print("="*60)
    cpu_count = os.cpu_count() or 4
    print(f"CPU Cores Available:       {cpu_count}")
    print(f"Upload Workers:            {config.max_upload_workers} (concurrent uploads)")
    print(f"Hash Workers:              {config.max_hash_workers} (parallel hashing)")
    print(f"Markdown Workers:          {config.max_markdown_workers} (parallel conversion)")
    print(f"Batch Metadata Updates:    Enabled (20 items/batch)")
    print("="*60 + "\n")

    # Show sync mode to user
    if config.force_upload:
        print("[!] Force upload mode enabled - all files will be uploaded regardless of changes")
    else:
        print("[OK] Smart sync mode enabled - unchanged files will be skipped")

    # Show markdown conversion mode
    if config.convert_md_to_html:
        print("[OK] Markdown to HTML conversion enabled - .md files will be converted with Mermaid diagrams as SVG")
    else:
        print("[!] Markdown to HTML conversion disabled - .md files will be uploaded as-is")

    # Show exclusion patterns if any
    if config.exclude_patterns_list:
        print(f"[=] Exclusion patterns enabled: {', '.join(config.exclude_patterns_list)}")

    # Show sync deletion mode
    if config.sync_delete:
        if config.sync_delete_whatif:
            print("[!] Sync deletion enabled in WHATIF mode - will show what would be deleted without actually deleting")
        else:
            print("[!] Sync deletion enabled - files in SharePoint but not in sync set will be DELETED")
    else:
        print("[OK] Sync deletion disabled - no files will be removed from SharePoint")

    # Show file discovery details
    print(f"[=] Current working directory: {os.getcwd()}")
    print(f"[=] Searching for files matching pattern: {config.file_path}")
    print(f"[=] Recursive search: {config.recursive}")

    # Discover files based on glob pattern and exclusions
    local_files, local_dirs = discover_files(
        config.file_path,
        config.recursive,
        config.exclude_patterns_list
    )

    if not local_files:
        print("[!] No files matched the pattern")
        sys.exit(1)

    print(f"[*] Found {len(local_files)} files to process")

    # Calculate base path for maintaining folder structure
    base_path = calculate_base_path(local_files, local_dirs)

    # Create Graph client and connect to SharePoint
    print("[*] Connecting to SharePoint...")
    try:
        client = create_graph_client(
            config.tenant_id, config.client_id, config.client_secret,
            config.login_endpoint, config.graph_endpoint
        )
        root_drive = client.sites.get_by_url(config.tenant_url).drive.root.get_by_path(config.upload_path)

        # Execute the query to test connection and permissions
        # This also initializes the root_drive object for use
        root_drive.get().execute_query()
        print(f"[✓] Connected to SharePoint at: {config.upload_path}")

        # Check and create FileHash column if needed
        library_name = get_library_name_from_path(config.upload_path)

        # Attempt to create the FileHash column for hash-based comparison
        filehash_column_available, library_name = check_and_create_filehash_column(
            config.tenant_url, library_name,
            config.tenant_id, config.client_id, config.client_secret,
            config.login_endpoint, config.graph_endpoint
        )
        if filehash_column_available:
            print("[✓] FileHash column is available for hash-based comparison")
        else:
            print("[!] FileHash column not available, will use size-based comparison")

    except Exception as conn_error:
        # Connection failed - provide helpful troubleshooting info
        print(f"[Error] Failed to connect to SharePoint: {conn_error}")
        print("[!] Ensure that:")
        print("    - Your credentials are correct")
        print("    - The site URL is correct")
        print("    - The upload path exists on the SharePoint site")
        print("    - You have appropriate permissions")
        sys.exit(1)  # Exit with error code

    # Parallel upload - process all files concurrently
    # Track converted files to avoid uploading .md files when .html versions exist
    converted_md_files = set()

    # Show parallel processing info
    print(f"[OK] Using parallel processing (Upload workers: {config.max_upload_workers}, Hash workers: {config.max_hash_workers})")

    # Initialize parallel uploader with auto-detected workers
    parallel_uploader = ParallelUploader(
        max_workers=config.max_upload_workers,
        upload_stats_instance=upload_stats,
        batch_metadata_updates=True  # Always use batch metadata updates
    )

    # Process all files in parallel
    failed_count = parallel_uploader.process_files(
        local_files,
        root_drive,
        base_path,
        config,
        filehash_column_available,
        library_name,
        converted_md_files
    )

    # Perform sync deletion if enabled
    if config.sync_delete:
        perform_sync_deletion(root_drive, local_files, base_path, config)

    # Print final summary report
    # Pass whatif mode status for proper deletion statistics labeling
    whatif_mode = config.sync_delete and config.sync_delete_whatif
    print_summary(len(local_files), whatif_mode=whatif_mode)

    # Exit with appropriate code
    # Exit code 0 = success, 1 = failure
    total_failed = failed_count + upload_stats.stats['failed_files']
    if total_failed > 0:
        # Some files failed - signal error to CI system
        print(f"[!] {total_failed} file(s) failed to process")
        sys.exit(1)

    # If we get here, all uploads succeeded (exit code 0 is implicit)
    if is_debug_enabled():
        print("[✓] All files processed successfully")


if __name__ == "__main__":
    main()
