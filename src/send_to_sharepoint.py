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
    python send_to_sharepoint.py <site_name> <sharepoint_host> <tenant_id>
                                 <client_id> <client_secret> <upload_path>
                                 <file_path> [max_retry] [login_endpoint]
                                 [graph_endpoint] [recursive] [force_upload]
                                 [convert_md_to_html]

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

# Standard library imports (come with Python)
import sys        # Provides access to command-line arguments and exit codes
import os         # Operating system interface for file/directory operations
import glob       # Unix-style pathname pattern expansion (e.g., *.txt matches all .txt files)
import time       # Time-related functions for delays and retries
import tempfile   # Temporary file and directory creation
import shutil     # High-level file operations (copy, move, etc.)
import xxhash     # For fast xxHash128 non-cryptographic hashing
from datetime import datetime  # For HTML generation date/time display
import requests   # For direct Graph API calls

# Third-party library imports (need to be installed via pip)
from dotenv import load_dotenv  # Load environment variables from .env file
import msal       # Microsoft Authentication Library for Azure AD authentication
import mistune   # Fast markdown parser for converting MD to HTML
import subprocess # For running mermaid-cli to convert diagrams to SVG
import re        # Regular expressions for pattern matching

# Load environment variables from .env file if it exists
# This allows local development and Docker to use consistent configuration
load_dotenv()

# Office365 library imports for SharePoint/OneDrive interaction
from office365.graph_client import GraphClient  # Main client for Microsoft Graph API
from office365.runtime.odata.v4.upload_session_request import UploadSessionRequest  # For large file uploads
from office365.onedrive.driveitems.driveItem import DriveItem  # Represents files/folders in OneDrive
from office365.onedrive.internal.paths.url import UrlPath  # URL path utilities
from office365.runtime.queries.upload_session import UploadSessionQuery  # Upload session management
from office365.onedrive.driveitems.uploadable_properties import DriveItemUploadableProperties  # File properties

# ====================================================================
# COMMAND-LINE ARGUMENTS PARSING
# ====================================================================
# sys.argv is a list containing command-line arguments
# sys.argv[0] is the script name itself, so we start from index 1
# Example: python script.py arg1 arg2 → sys.argv = ['script.py', 'arg1', 'arg2']

# Required arguments (script will fail if these are missing)
site_name = sys.argv[1]              # SharePoint site name (e.g., 'TeamDocuments')
sharepoint_host_name = sys.argv[2]   # SharePoint domain (e.g., 'company.sharepoint.com')
tenant_id = sys.argv[3]              # Azure AD tenant ID (GUID format)
client_id = sys.argv[4]              # App registration client ID
client_secret = sys.argv[5]          # App registration client secret (keep secure!)
upload_path = sys.argv[6]            # Target folder in SharePoint (e.g., 'Documents/Reports')
file_path = sys.argv[7]              # Local file/folder to upload (supports wildcards like *.pdf)

# Optional arguments with default values
# The 'or' operator returns the right value if left value is falsy (0, None, empty string)
max_retry = int(sys.argv[8]) or 3    # Number of upload retries (default: 3)

# Use default endpoints if not provided (for special Azure environments like GovCloud)
login_endpoint = sys.argv[9] or "login.microsoftonline.com"    # Azure AD login URL
graph_endpoint = sys.argv[10] or "graph.microsoft.com"         # Microsoft Graph API URL

# Check if recursive flag is provided (for searching subdirectories)
# len(sys.argv) > 11 ensures we don't get IndexError if argument doesn't exist
file_path_recursive_match = sys.argv[11] if len(sys.argv) > 11 and sys.argv[11] else "False"

# Check if force upload flag is provided (skip file comparison)
# len(sys.argv) > 12 ensures we don't get IndexError if argument doesn't exist
force_upload_flag = sys.argv[12] if len(sys.argv) > 12 and sys.argv[12] else "False"

# Check if markdown to HTML conversion flag is provided
# len(sys.argv) > 13 ensures we don't get IndexError if argument doesn't exist
convert_md_to_html_flag = sys.argv[13] if len(sys.argv) > 13 and sys.argv[13] else "True"

# ====================================================================
# SHAREPOINT FILENAME SANITIZATION
# ====================================================================

def sanitize_sharepoint_name(name, is_folder=False):
    """
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

# ====================================================================
# FILE HASHING WITH XXHASH128 FOR FAST COMPARISON
# ====================================================================

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

# ====================================================================
# MARKDOWN TO HTML CONVERSION WITH MERMAID SUPPORT
# ====================================================================

def convert_mermaid_to_svg(mermaid_code):
    """
    Convert Mermaid diagram code to SVG using mermaid-cli.

    Uses the mmdc command-line tool installed via npm to render
    Mermaid diagrams as static SVG images.

    Args:
        mermaid_code (str): Mermaid diagram definition

    Returns:
        str: SVG content as string, or None if conversion fails
    """
    try:
        # Create temporary files for input and output
        with tempfile.NamedTemporaryFile(mode='w', suffix='.mmd', delete=False) as mmd_file:
            mmd_file.write(mermaid_code)
            mmd_path = mmd_file.name

        svg_path = mmd_path.replace('.mmd', '.svg')

        # Run mermaid-cli to convert to SVG
        # Using puppeteer config to work in Docker container
        result = subprocess.run(
            ['mmdc', '-i', mmd_path, '-o', svg_path, '--puppeteerConfigFile', '/usr/src/app/puppeteer-config.json'],
            capture_output=True,
            text=True,
            timeout=30
        )

        if result.returncode == 0 and os.path.exists(svg_path):
            # Read the generated SVG
            with open(svg_path, 'r', encoding='utf-8') as f:
                svg_content = f.read()

            # Clean up temporary files
            os.unlink(mmd_path)
            os.unlink(svg_path)

            return svg_content
        else:
            print(f"[!] Mermaid conversion failed: {result.stderr}")
            # Clean up temp file
            if os.path.exists(mmd_path):
                os.unlink(mmd_path)
            if os.path.exists(svg_path):
                os.unlink(svg_path)
            return None

    except Exception as e:
        print(f"[!] Error converting Mermaid diagram: {e}")
        return None

def convert_markdown_to_html(md_content, filename):
    """
    Convert Markdown content to HTML with Mermaid diagrams rendered as SVG.

    This function:
    1. Parses markdown using Mistune
    2. Finds and converts Mermaid code blocks to inline SVG
    3. Applies GitHub-like styling for SharePoint viewing

    Args:
        md_content (str): Markdown content to convert
        filename (str): Original filename for the HTML title

    Returns:
        str: Complete HTML document with embedded styles and SVGs
    """
    # First, extract and convert all mermaid blocks to placeholder SVGs
    mermaid_pattern = r'```mermaid\n(.*?)\n```'
    mermaid_blocks = []

    def replace_mermaid_with_placeholder(match):
        mermaid_code = match.group(1)
        placeholder = f"<!--MERMAID_PLACEHOLDER_{len(mermaid_blocks)}-->"

        # Convert to SVG
        svg_content = convert_mermaid_to_svg(mermaid_code)
        if svg_content:
            # Clean up the SVG for inline embedding
            # Remove XML declaration if present
            svg_content = re.sub(r'<\?xml[^>]*\?>', '', svg_content)
            svg_content = svg_content.strip()
            mermaid_blocks.append(svg_content)
        else:
            # If conversion failed, keep as code block
            mermaid_blocks.append(f'<pre><code>mermaid\n{mermaid_code}</code></pre>')

        return placeholder

    # Replace mermaid blocks with placeholders
    md_with_placeholders = re.sub(mermaid_pattern, replace_mermaid_with_placeholder, md_content, flags=re.DOTALL)

    # Convert markdown to HTML using Mistune
    html_body = mistune.html(md_with_placeholders)

    # Replace placeholders with actual SVG content
    for i, svg_content in enumerate(mermaid_blocks):
        placeholder = f"<!--MERMAID_PLACEHOLDER_{i}-->"
        # Wrap SVG in a div for centering
        wrapped_svg = f'<div class="mermaid-diagram">{svg_content}</div>'
        html_body = html_body.replace(f"<p>{placeholder}</p>", wrapped_svg)
        html_body = html_body.replace(placeholder, wrapped_svg)

    # Create the complete HTML document
    html_template = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>{filename.replace('.md', '')}</title>

    <style>
        /* GitHub-like styling for SharePoint */
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Noto Sans", Helvetica, Arial, sans-serif;
            font-size: 16px;
            line-height: 1.5;
            word-wrap: break-word;
            padding: 20px;
            max-width: 980px;
            margin: 0 auto;
            color: #1F2328;
            background: #ffffff;
        }}

        h1, h2, h3, h4, h5, h6 {{
            margin-top: 24px;
            margin-bottom: 16px;
            font-weight: 600;
            line-height: 1.25;
        }}

        h1 {{
            font-size: 2em;
            border-bottom: 1px solid #d1d9e0;
            padding-bottom: .3em;
        }}

        h2 {{
            font-size: 1.5em;
            border-bottom: 1px solid #d1d9e0;
            padding-bottom: .3em;
        }}

        h3 {{ font-size: 1.25em; }}
        h4 {{ font-size: 1em; }}
        h5 {{ font-size: .875em; }}
        h6 {{ font-size: .85em; color: #59636e; }}

        code {{
            padding: .2em .4em;
            margin: 0;
            font-size: 85%;
            white-space: break-spaces;
            background-color: #f6f8fa;
            border-radius: 6px;
            font-family: ui-monospace, SFMono-Regular, "SF Mono", Consolas, "Liberation Mono", Menlo, monospace;
        }}

        pre {{
            padding: 16px;
            overflow: auto;
            font-size: 85%;
            line-height: 1.45;
            color: #1F2328;
            background-color: #f6f8fa;
            border-radius: 6px;
            margin-top: 0;
            margin-bottom: 16px;
        }}

        pre code {{
            display: inline;
            max-width: auto;
            padding: 0;
            margin: 0;
            overflow: visible;
            line-height: inherit;
            word-wrap: normal;
            background-color: transparent;
            border: 0;
        }}

        blockquote {{
            margin: 0;
            padding: 0 1em;
            color: #59636e;
            border-left: .25em solid #d1d9e0;
        }}

        table {{
            border-spacing: 0;
            border-collapse: collapse;
            display: block;
            width: max-content;
            max-width: 100%;
            overflow: auto;
            margin-top: 0;
            margin-bottom: 16px;
        }}

        table th {{
            font-weight: 600;
            padding: 6px 13px;
            border: 1px solid #d1d9e0;
            background-color: #f6f8fa;
        }}

        table td {{
            padding: 6px 13px;
            border: 1px solid #d1d9e0;
        }}

        table tr:nth-child(2n) {{
            background-color: #f6f8fa;
        }}

        ul, ol {{
            margin-top: 0;
            margin-bottom: 16px;
            padding-left: 2em;
        }}

        ul ul, ul ol, ol ol, ol ul {{
            margin-top: 0;
            margin-bottom: 0;
        }}

        li > p {{
            margin-top: 16px;
        }}

        a {{
            color: #0969da;
            text-decoration: none;
        }}

        a:hover {{
            text-decoration: underline;
        }}

        hr {{
            height: .25em;
            padding: 0;
            margin: 24px 0;
            background-color: #d1d9e0;
            border: 0;
        }}

        img {{
            max-width: 100%;
            box-sizing: content-box;
        }}

        /* Mermaid diagram container */
        .mermaid-diagram {{
            text-align: center;
            margin: 16px 0;
            padding: 16px;
            background-color: #f6f8fa;
            border-radius: 6px;
            overflow-x: auto;
        }}

        .mermaid-diagram svg {{
            max-width: 100%;
            height: auto;
        }}

        /* Task list items */
        .task-list-item {{
            list-style-type: none;
        }}

        .task-list-item input {{
            margin: 0 .2em .25em -1.4em;
            vertical-align: middle;
        }}

    </style>
</head>
<body>
    {html_body}
</body>
</html>'''

    return html_template

# ====================================================================
# URL AND CONFIGURATION SETUP
# ====================================================================

# Construct the full SharePoint site URL
# f-strings (f"...") allow embedding variables directly in strings
tenant_url = f'https://{sharepoint_host_name}/sites/{site_name}'

# Convert string arguments to boolean
# .lower() converts to lowercase for case-insensitive comparison
recursive = file_path_recursive_match.lower() in ['true', '1', 'yes']
force_upload = force_upload_flag.lower() in ['true', '1', 'yes']
convert_md_to_html = convert_md_to_html_flag.lower() in ['true', '1', 'yes']

# Show sync mode to user
if force_upload:
    print("[!] Force upload mode enabled - all files will be uploaded regardless of changes")
else:
    print("[OK] Smart sync mode enabled - unchanged files will be skipped")

# Show markdown conversion mode
if convert_md_to_html:
    print("[OK] Markdown to HTML conversion enabled - .md files will be converted with Mermaid diagrams as SVG")
else:
    print("[!] Markdown to HTML conversion disabled - .md files will be uploaded as-is")

# ====================================================================
# FILE DISCOVERY - Finding files to upload
# ====================================================================

# Use glob to find all files/directories matching the pattern
# glob.glob() returns a list of paths matching a pathname pattern
# Examples: '*.txt' finds all .txt files, '**/*.py' finds all .py files recursively
local_items = glob.glob(file_path, recursive=recursive)

# Exit with error if no matches found
if not local_items:
    print(f"[Error] No files or directories matched pattern: {file_path}")
    sys.exit(1)  # Exit code 1 indicates error to calling process (e.g., GitHub Actions)

# ====================================================================
# FILE AND DIRECTORY SEPARATION
# ====================================================================

# Initialize empty lists to store files and directories separately
local_files = []  # Will contain paths to actual files
local_dirs = []   # Will contain paths to directories

# Iterate through each matched item and categorize it
for item in local_items:
    if os.path.isfile(item):  # Check if path points to a file
        local_files.append(item)  # Add to files list
    elif os.path.isdir(item):  # Check if path points to a directory
        local_dirs.append(item)   # Add to directories list

        # For directories, we need to get all files inside them
        # os.walk() recursively traverses directory tree
        # It yields (current_dir, subdirectories, files) for each directory
        for root, dirs, files in os.walk(item):
            for file in files:
                # os.path.join() creates proper path regardless of OS (Windows/Mac/Linux)
                # Windows uses backslash (\), Unix uses forward slash (/)
                local_files.append(os.path.join(root, file))

# Final validation - ensure we have something to upload
if not local_files and not local_dirs:
    print(f"[Error] No files or directories found matching pattern: {file_path}")
    sys.exit(1)

# Inform user about what was found
print(f"Found {len(local_files)} file(s) and {len(local_dirs)} directory(ies) to process")

def acquire_token():
    """
    Acquire an authentication token from Azure Active Directory using MSAL.

    This function handles the OAuth 2.0 client credentials flow, which is used
    for service-to-service authentication (no user interaction required).

    Returns:
        dict: Token dictionary containing:
            - 'access_token': The JWT token to authenticate API calls
            - 'token_type': Usually 'Bearer'
            - 'expires_in': Token lifetime in seconds

    Raises:
        Exception: If authentication fails (wrong credentials, network issues, etc.)

    Example:
        token = acquire_token()
        headers = {'Authorization': f"{token['token_type']} {token['access_token']}"}

    Note:
        This uses the client credentials flow, suitable for automated scripts.
        The app registration must have Graph API Sites.ReadWrite.All permission.
    """
    # Build the Azure AD authority URL
    # Format: https://login.microsoftonline.com/{tenant_id}
    authority_url = f'https://{login_endpoint}/{tenant_id}'

    # Create MSAL confidential client application
    # 'Confidential' means it can securely store credentials (unlike public/mobile apps)
    app = msal.ConfidentialClientApplication(
        authority=authority_url,           # Azure AD endpoint
        client_id=client_id,              # Your app registration's ID
        client_credential=client_secret    # Your app's secret key
    )

    # Request an access token for Microsoft Graph API
    # '/.default' scope means "use all permissions granted to this app"
    token = app.acquire_token_for_client(scopes=[f"https://{graph_endpoint}/.default"])

    return token

def check_and_create_filehash_column(site_url, list_name):
    """
    Check if FileHash column exists in SharePoint document library and create if needed.

    Uses direct Graph API calls to bypass Office365-REST-Python-Client limitations.
    This ensures the FileHash column is available for storing file hashes.

    Args:
        site_url (str): Full SharePoint site URL
        list_name (str): Name of the document library (usually "Documents")

    Returns:
        bool: True if column exists or was created successfully, False otherwise

    Note:
        Requires Sites.ReadWrite.All or Sites.Manage.All permissions.
        The column is created as a single line of text with 32 character limit
        (exact length of xxHash128 hexadecimal representation).
    """
    try:
        # Get token for Graph API
        print("[?] Checking for FileHash column in SharePoint...")
        token = acquire_token()

        if 'access_token' not in token:
            print(f"[!] Failed to acquire token for Graph API: {token.get('error_description', 'Unknown error')}")
            return False

        headers = {
            'Authorization': f"Bearer {token['access_token']}",
            'Content-Type': 'application/json'
        }

        # Parse site URL to get site ID
        # Format: https://tenant.sharepoint.com/sites/sitename
        site_parts = site_url.replace('https://', '').split('/')
        host_name = site_parts[0]
        site_name = site_parts[2] if len(site_parts) > 2 else ''

        # Get site ID first
        site_endpoint = f"https://{graph_endpoint}/v1.0/sites/{host_name}:/sites/{site_name}"
        site_response = requests.get(site_endpoint, headers=headers)

        if site_response.status_code != 200:
            print(f"[!] Failed to get site information: {site_response.status_code}")
            print(f"[DEBUG] Response: {site_response.text[:500]}")
            return False

        site_data = site_response.json()
        site_id = site_data.get('id')

        if not site_id:
            print("[!] Could not retrieve site ID")
            return False

        # Get the document library (list) ID
        lists_endpoint = f"https://{graph_endpoint}/v1.0/sites/{site_id}/lists"
        lists_response = requests.get(lists_endpoint, headers=headers)

        if lists_response.status_code != 200:
            print(f"[!] Failed to get lists: {lists_response.status_code}")
            return False

        lists_data = lists_response.json()
        list_id = None

        # Find the document library by name
        for lst in lists_data.get('value', []):
            if lst.get('displayName') == list_name or lst.get('name') == list_name:
                list_id = lst.get('id')
                break

        if not list_id:
            # Try "Shared Documents" as fallback
            for lst in lists_data.get('value', []):
                if lst.get('displayName') == 'Shared Documents' or lst.get('name') == 'Shared Documents':
                    list_id = lst.get('id')
                    print(f"[!] Using 'Shared Documents' instead of '{list_name}'")
                    break

        if not list_id:
            print(f"[!] Document library '{list_name}' not found")
            return False

        # Check if FileHash column already exists
        columns_endpoint = f"https://{graph_endpoint}/v1.0/sites/{site_id}/lists/{list_id}/columns"
        columns_response = requests.get(columns_endpoint, headers=headers)

        if columns_response.status_code != 200:
            print(f"[!] Failed to get columns: {columns_response.status_code}")
            return False

        columns_data = columns_response.json()
        filehash_exists = False

        # Check for existing FileHash column
        for column in columns_data.get('value', []):
            if column.get('name') == 'FileHash' or column.get('displayName') == 'FileHash':
                filehash_exists = True
                print("[✓] FileHash column already exists")
                break

        # Create column if it doesn't exist
        if not filehash_exists:
            print("[+] Creating FileHash column...")

            # Column definition for FileHash
            column_definition = {
                "displayName": "FileHash",
                "name": "FileHash",
                "description": "xxHash128 checksum for file content verification",
                "enforceUniqueValues": False,
                "hidden": False,
                "indexed": False,
                "readOnly": False,
                "required": False,
                "text": {
                    "allowMultipleLines": False,
                    "appendChangesToExistingText": False,
                    "linesForEditing": 0,
                    "maxLength": 32  # xxHash128 produces 32-character hex string
                }
            }

            # Create the column
            create_response = requests.post(
                columns_endpoint,
                headers=headers,
                json=column_definition
            )

            if create_response.status_code == 201:
                print("[✓] FileHash column created successfully")
                return True
            else:
                print(f"[!] Failed to create FileHash column: {create_response.status_code}")
                print(f"[DEBUG] Response: {create_response.text[:500]}")
                return False

        return True

    except Exception as e:
        print(f"[!] Error checking/creating FileHash column: {e}")
        return False

def rewrite_endpoint(request):
    """
    Modify API request URLs for non-standard Microsoft Graph endpoints.

    This function is needed for special Azure environments like:
    - Azure Government Cloud (graph.microsoft.us)
    - Azure Germany (graph.microsoft.de)
    - Azure China (microsoftgraph.chinacloudapi.cn)

    Args:
        request: The HTTP request object to be modified

    Note:
        This is a callback function used by the GraphClient to intercept
        and modify requests before they're sent.
    """
    # Replace default endpoint with custom one if specified
    request.url = request.url.replace(
        "https://graph.microsoft.com", f"https://{graph_endpoint}"
    )

def get_sharepoint_list_item_by_filename(site_url, list_name, filename):
    """
    Get SharePoint list item by filename using direct Graph API REST calls.

    Args:
        site_url (str): Full SharePoint site URL
        list_name (str): Name of the document library (usually "Documents")
        filename (str): Name of the file to find

    Returns:
        dict: List item data with custom columns, or None if not found
    """
    try:
        # Get token for Graph API
        token = acquire_token()

        if 'access_token' not in token:
            print(f"[!] Failed to acquire token for Graph API: {token.get('error_description', 'Unknown error')}")
            return None

        headers = {
            'Authorization': f"Bearer {token['access_token']}",
            'Content-Type': 'application/json'
        }

        # Parse site URL to get site ID
        site_parts = site_url.replace('https://', '').split('/')
        host_name = site_parts[0]
        site_name = site_parts[2] if len(site_parts) > 2 else ''

        # Get site ID first
        site_endpoint = f"https://{graph_endpoint}/v1.0/sites/{host_name}:/sites/{site_name}"
        site_response = requests.get(site_endpoint, headers=headers)

        if site_response.status_code != 200:
            print(f"[!] Failed to get site information: {site_response.status_code}")
            return None

        site_data = site_response.json()
        site_id = site_data.get('id')

        if not site_id:
            print("[!] Could not retrieve site ID")
            return None

        # Get the document library (list) ID
        lists_endpoint = f"https://{graph_endpoint}/v1.0/sites/{site_id}/lists"
        lists_response = requests.get(lists_endpoint, headers=headers)

        if lists_response.status_code != 200:
            print(f"[!] Failed to get lists: {lists_response.status_code}")
            return None

        lists_data = lists_response.json()
        list_id = None

        for sp_list in lists_data.get('value', []):
            if sp_list.get('displayName') == list_name or sp_list.get('name') == list_name:
                list_id = sp_list.get('id')
                break

        if not list_id:
            print(f"[!] Could not find list '{list_name}'")
            return None

        # Query list items by filename with expanded fields
        # Filter by FileLeafRef which contains the filename
        items_endpoint = f"https://{graph_endpoint}/v1.0/sites/{site_id}/lists/{list_id}/items"
        items_params = {
            '$expand': 'fields',
            '$filter': f"fields/FileLeafRef eq '{filename}'"
        }

        items_response = requests.get(items_endpoint, headers=headers, params=items_params)

        if items_response.status_code != 200:
            print(f"[!] Failed to get list items: {items_response.status_code}")
            return None

        items_data = items_response.json()
        items = items_data.get('value', [])

        if items:
            return items[0]  # Return first matching item
        else:
            return None

    except Exception as e:
        print(f"[!] Error getting list item by filename: {str(e)[:200]}")
        return None

def update_sharepoint_list_item_field(site_url, list_name, item_id, field_name, field_value):
    """
    Update a custom field in a SharePoint list item using direct Graph API REST calls.

    Args:
        site_url (str): Full SharePoint site URL
        list_name (str): Name of the document library (usually "Documents")
        item_id (str): SharePoint list item ID
        field_name (str): Internal name of the field to update
        field_value (str): Value to set for the field

    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Get token for Graph API
        token = acquire_token()

        if 'access_token' not in token:
            print(f"[!] Failed to acquire token for Graph API: {token.get('error_description', 'Unknown error')}")
            return False

        headers = {
            'Authorization': f"Bearer {token['access_token']}",
            'Content-Type': 'application/json'
        }

        # Parse site URL to get site ID
        site_parts = site_url.replace('https://', '').split('/')
        host_name = site_parts[0]
        site_name = site_parts[2] if len(site_parts) > 2 else ''

        # Get site ID first
        site_endpoint = f"https://{graph_endpoint}/v1.0/sites/{host_name}:/sites/{site_name}"
        site_response = requests.get(site_endpoint, headers=headers)

        if site_response.status_code != 200:
            print(f"[!] Failed to get site information: {site_response.status_code}")
            return False

        site_data = site_response.json()
        site_id = site_data.get('id')

        if not site_id:
            print("[!] Could not retrieve site ID")
            return False

        # Get the document library (list) ID
        lists_endpoint = f"https://{graph_endpoint}/v1.0/sites/{site_id}/lists"
        lists_response = requests.get(lists_endpoint, headers=headers)

        if lists_response.status_code != 200:
            print(f"[!] Failed to get lists: {lists_response.status_code}")
            return False

        lists_data = lists_response.json()
        list_id = None

        for sp_list in lists_data.get('value', []):
            if sp_list.get('displayName') == list_name or sp_list.get('name') == list_name:
                list_id = sp_list.get('id')
                break

        if not list_id:
            print(f"[!] Could not find list '{list_name}'")
            return False

        # Update the field using PATCH request
        fields_endpoint = f"https://{graph_endpoint}/v1.0/sites/{site_id}/lists/{list_id}/items/{item_id}/fields"
        field_data = {field_name: field_value}

        update_response = requests.patch(fields_endpoint, headers=headers, json=field_data)

        if update_response.status_code == 200:
            return True
        else:
            print(f"[!] Failed to update field: {update_response.status_code}")
            print(f"[DEBUG] Response: {update_response.text[:500]}")
            return False

    except Exception as e:
        print(f"[!] Error updating list item field: {str(e)[:200]}")
        return False

# ====================================================================
# MICROSOFT GRAPH CLIENT SETUP
# ====================================================================

# Initialize the Graph API client with our authentication function
# GraphClient will call acquire_token() whenever it needs a fresh token
client = GraphClient(acquire_token)

# Register our endpoint rewriter to handle non-standard environments
# The 'False' parameter means don't execute immediately
client.before_execute(rewrite_endpoint, False)

# ====================================================================
# SHAREPOINT CONNECTION AND SETUP
# ====================================================================

# Get the target folder in SharePoint where files will be uploaded
# This chains multiple API calls:
# 1. sites.get_by_url() - Gets the SharePoint site
# 2. .drive - Gets the site's default document library
# 3. .root - Gets the root folder of the drive
# 4. .get_by_path() - Navigates to our target upload folder
root_drive = client.sites.get_by_url(tenant_url).drive.root.get_by_path(upload_path)

# ====================================================================
# GLOBAL STATE TRACKING
# ====================================================================

# Dictionary to cache created folders (avoids redundant API calls)
# Key: folder path, Value: DriveItem object
created_folders = {}

# Statistics tracker for upload summary
# Using a dictionary makes it easy to pass by reference and update from functions
upload_stats = {
    'new_files': 0,       # Files that didn't exist in SharePoint
    'replaced_files': 0,  # Files that were overwritten
    'skipped_files': 0,   # Files skipped because they're identical
    'failed_files': 0,    # Files that failed to upload
    'bytes_uploaded': 0,  # Total bytes uploaded
    'bytes_skipped': 0    # Total bytes skipped
}

def ensure_folder_exists(parent_drive, folder_path):
    """
    Recursively create folder structure in SharePoint if it doesn't exist.

    This function handles nested folder creation, ensuring the entire path
    exists before uploading files. It uses caching to avoid redundant API calls.

    Args:
        parent_drive (DriveItem): The parent folder where structure should be created
        folder_path (str): Path to create (e.g., 'folder1/folder2/folder3')

    Returns:
        DriveItem: The final folder in the path, ready to receive files

    Raises:
        Exception: If folder creation fails after all retry attempts

    Example:
        # Create nested folders
        target = ensure_folder_exists(root_drive, "2024/Reports/January")
        # Now upload file to the January folder
        upload_file(target, "report.pdf", chunk_size)

    Note:
        - Caches created folders to minimize API calls
        - Handles both forward slash (/) and backslash (\) path separators
        - Sanitizes folder names for SharePoint compatibility
        - Compatible with Office365-REST-Python-Client v2.6.2+
    """
    # Convert Windows backslashes to forward slashes for consistency
    # This ensures the function works on both Windows and Unix systems
    folder_path = folder_path.replace('\\', '/')

    # Sanitize the entire path for SharePoint compatibility
    original_folder_path = folder_path
    folder_path = sanitize_path_components(folder_path)

    # Check cache first to avoid unnecessary API calls
    # This significantly improves performance for large folder structures
    if folder_path in created_folders:
        return created_folders[folder_path]

    # Split path into individual folder names
    # List comprehension filters out empty strings from split()
    # Example: "a/b/c" becomes ['a', 'b', 'c']
    path_parts = [part for part in folder_path.split('/') if part]

    # If no folders to create, return the parent
    if not path_parts:
        return parent_drive

    # Start from the parent folder
    current_drive = parent_drive
    current_path = ""  # Track the path we've built so far

    # Process each folder in the path
    for folder_name in path_parts:
        # Note: folder_name is already sanitized from path_parts split
        # Build cumulative path as we go deeper
        # Ternary operator: use "/" separator if current_path exists, else start fresh
        current_path = f"{current_path}/{folder_name}" if current_path else folder_name

        # Skip if we've already processed this folder path
        if current_path in created_folders:
            current_drive = created_folders[current_path]
            continue  # Move to next folder in path

        # ============================================================
        # STEP 1: Check if folder already exists in SharePoint
        # ============================================================
        folder_found = False  # Flag to track if folder exists

        try:
            print(f"[?] Checking if folder exists: {current_path}")

            # Get all items (files and folders) in current folder
            # execute_query() sends the API request and waits for response
            children = current_drive.children.get().execute_query()

            # Iterate through children to find matching folder
            for child in children:
                # Check two conditions:
                # 1. Name matches what we're looking for
                # 2. It's a folder (has 'folder' attribute), not a file
                if child.name == folder_name and hasattr(child, 'folder'):
                    # Folder found! Update references and cache it
                    current_drive = child
                    created_folders[current_path] = child
                    print(f"[✓] Folder already exists: {current_path}")
                    folder_found = True
                    break  # Stop searching once found

        except Exception as e:
            # API call failed - assume folder doesn't exist
            print(f"[!] Error checking folder existence: {e}")
            folder_found = False

        # ============================================================
        # STEP 2: Create folder if it doesn't exist
        # ============================================================
        if not folder_found:
            try:
                print(f"[+] Creating folder: {folder_name}")

                # For Office365-REST-Python-Client v2.6.2, use create_folder method
                # This is a built-in method specifically for folder creation
                created_folder = current_drive.create_folder(folder_name).execute_query()

                current_drive = created_folder
                created_folders[current_path] = created_folder
                print(f"[✓] Created folder: {current_path}")

            except AttributeError:
                # If create_folder method doesn't exist, try alternative approach
                try:
                    print(f"[!] create_folder not available, trying add() method")

                    # Create a DriveItem instance for the folder
                    new_folder = DriveItem(current_drive.context)
                    new_folder.set_property("name", folder_name)
                    new_folder.set_property("folder", {})

                    # Use add() without parameters (the object already has properties set)
                    created_folder = current_drive.children.add()
                    created_folder.set_property("name", folder_name)
                    created_folder.set_property("folder", {})
                    created_folder.execute_query()

                    current_drive = created_folder
                    created_folders[current_path] = created_folder
                    print(f"[✓] Created folder: {current_path}")

                except Exception as add_error:
                    print(f"[!] add() method failed: {add_error}")
                    raise

            except Exception as create_error:
                error_msg = str(create_error)

                # Check if folder already exists (common race condition)
                if "nameAlreadyExists" in error_msg or "already exists" in error_msg.lower():
                    print(f"[!] Folder already exists (race condition): {folder_name}")
                    try:
                        # Try to get the existing folder
                        children = current_drive.children.get().execute_query()
                        for child in children:
                            if child.name == folder_name and hasattr(child, 'folder'):
                                current_drive = child
                                created_folders[current_path] = child
                                print(f"[✓] Found existing folder: {current_path}")
                                break
                    except:
                        pass
                else:
                    print(f"[!] Error creating folder {folder_name}: {create_error}")

                    # Fallback: Try to navigate to folder in case it exists
                    try:
                        print(f"[!] Attempting to navigate to folder: {folder_name}")
                        test_folder = current_drive.get_by_path(folder_name).get().execute_query()
                        if test_folder and hasattr(test_folder, 'folder'):
                            current_drive = test_folder
                            created_folders[current_path] = test_folder
                            print(f"[✓] Successfully navigated to folder: {current_path}")
                        else:
                            raise Exception("Not a folder")
                    except Exception as nav_error:
                        print(f"[!] Unable to create or navigate to folder {current_path}")
                        print(f"[!] Will continue with parent folder")
                        # Don't fail the entire upload, just use parent folder
                        pass

    return current_drive

def progress_status(offset, file_size):
    print(f"Uploaded {offset} bytes from {file_size} bytes ... {offset/file_size*100:.2f}%")

def success_callback(remote_file, local_path, display_name=None):
    # Use display_name if provided (for temp files), otherwise use local_path
    file_display = display_name if display_name else local_path
    print(f"[✓] File {file_display} has been uploaded to {remote_file.web_url}")

def resumable_upload(drive, local_path, file_size, chunk_size, max_chunk_retry, timeout_secs):
    """
    Upload large files using resumable upload sessions.

    :param drive: The DriveItem representing the target folder
    :param local_path: Path to the local file to upload
    :param file_size: Size of the file in bytes
    :param chunk_size: Size of each chunk to upload
    :param max_chunk_retry: Maximum retries for each chunk
    :param timeout_secs: Total timeout in seconds
    """
    file_name = os.path.basename(local_path)
    # Sanitize the file name for SharePoint compatibility
    sanitized_name = sanitize_sharepoint_name(file_name, is_folder=False)

    # First, try the built-in upload_large_file method
    # This method handles the upload session creation properly
    try:
        print(f"[→] Using built-in upload method for large file: {sanitized_name}")
        if sanitized_name != file_name:
            print(f"    (Original name: {file_name})")
        with open(local_path, 'rb') as f:
            # Note: The built-in method might need the sanitized name set differently
            # We'll rely on the library to handle this correctly
            remote_file = drive.upload_large_file(f).execute_query()
            success_callback(remote_file, local_path, display_name=sanitized_name)
            return
    except AttributeError:
        # Method doesn't exist, continue with manual session
        print(f"[!] Built-in large file upload not available, using manual session")
    except Exception as e:
        print(f"[!] Built-in upload failed: {e}, trying manual session")

    # Manual upload session creation
    def _start_upload():
        with open(local_path, "rb") as local_file:
            session_request = UploadSessionRequest(
                local_file,
                chunk_size,
                lambda offset: progress_status(offset, file_size)
            )
            retry_seconds = timeout_secs / max_chunk_retry
            for session_request._range_data in session_request._read_next():
                for retry_number in range(max_chunk_retry):
                    try:
                        super(UploadSessionRequest, session_request).execute_query(qry)
                        break
                    except Exception as e:
                        if retry_number + 1 >= max_chunk_retry:
                            raise e
                        print(f"Retry {retry_number}: {e}")
                        time.sleep(retry_seconds)

    # Alternative approach: Use children.add() for upload session to avoid UrlPath issues
    # This approach works better for files with multiple periods in the name
    try:
        # Create a new DriveItem as a child of the folder
        return_type = drive.children.add()
        return_type.set_property("name", sanitized_name)
        return_type.set_property("file", {})

        # Create upload session query with conflict behavior
        upload_props = DriveItemUploadableProperties(name=sanitized_name)
        # Set conflict behavior to replace existing files
        upload_props.set_property("@microsoft.graph.conflictBehavior", "replace", False)

        qry = UploadSessionQuery(return_type, {"item": upload_props})
        drive.context.add_query(qry).after_query_execute(_start_upload)
        return_type.get().execute_query()
        success_callback(return_type, local_path)
    except Exception as e:
        print(f"[!] Children.add() approach failed: {e}")

        # Fallback: Try with the drive.upload method directly
        # This bypasses the manual upload session creation entirely
        print(f"[!] Attempting direct upload fallback for: {sanitized_name}")
        try:
            # For files > 4MB, we need to handle this differently
            # Let's try uploading in smaller chunks using the regular upload
            if file_size > 60*1024*1024:  # If > 60MB, fail
                raise Exception("File too large for fallback upload")

            # Create temporary file with sanitized name if needed
            temp_file_created = False
            upload_path = local_path

            if sanitized_name != file_name:
                temp_dir = tempfile.gettempdir()
                temp_path = os.path.join(temp_dir, sanitized_name)
                shutil.copy2(local_path, temp_path)
                upload_path = temp_path
                temp_file_created = True

            # Use regular upload as fallback
            remote_file = drive.upload_file(upload_path).execute_query()
            success_callback(remote_file, local_path, display_name=sanitized_name)

            # Clean up temp file if created
            if temp_file_created and os.path.exists(temp_path):
                os.remove(temp_path)

        except Exception as fallback_error:
            print(f"[!] Fallback upload also failed: {fallback_error}")

            # Last resort: Use the original UrlPath approach
            # (keeping for compatibility with older library versions)
            print(f"[!] Using original UrlPath approach as last resort")
            return_type = DriveItem(
                drive.context,
                UrlPath(sanitized_name, drive.resource_path))

            upload_props = DriveItemUploadableProperties(name=sanitized_name)
            upload_props.set_property("@microsoft.graph.conflictBehavior", "replace", False)

            qry = UploadSessionQuery(return_type, {"item": upload_props})
            drive.context.add_query(qry).after_query_execute(_start_upload)
            return_type.get().execute_query()
            success_callback(return_type, local_path, display_name=sanitized_name)

def check_file_needs_update(drive, local_path, file_name):
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

    Returns:
        tuple: (needs_update: bool, exists: bool, remote_file: DriveItem or None, local_hash: str or None)
            - needs_update: True if file should be uploaded
            - exists: True if file exists in SharePoint
            - remote_file: The existing SharePoint file object (if exists)
            - local_hash: The calculated hash of the local file (if computed)

    Example:
        needs_update, exists, remote, hash_val = check_file_needs_update(folder, "/path/to/file.pdf", "file.pdf")
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

    # Get debug flag
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
        remote_hash = None

        # Debug logging for FileHash retrieval
        debug_metadata = os.environ.get('DEBUG_METADATA', 'false').lower() == 'true'

        try:
            # Use direct Graph API REST calls to get SharePoint list item with custom columns
            list_item_data = get_sharepoint_list_item_by_filename(tenant_url, library_name, sanitized_name)

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
                        upload_stats['skipped_files'] += 1
                        upload_stats['bytes_skipped'] += local_size
                        return False, True, existing_file, local_hash
                    elif local_hash:
                        print(f"[*] File changed (hash mismatch): {sanitized_name}")
                        return True, True, existing_file, local_hash
                elif debug_metadata:
                    print(f"[DEBUG] FileHash not found in list item fields")
            elif debug_metadata:
                print(f"[DEBUG] Could not retrieve list item data for {sanitized_name}")

        except Exception as hash_error:
            # FileHash column might not exist or we can't access it
            print(f"[!] Could not retrieve FileHash via REST API, falling back to size comparison: {str(hash_error)[:100]}")
            hash_comparison_available = False

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
            debug_metadata = os.environ.get('DEBUG_METADATA', 'false').lower() == 'true'

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
                upload_stats['skipped_files'] += 1
                upload_stats['bytes_skipped'] += local_size
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

def check_and_delete_existing_file(drive, file_name):
    """
    Check if a file exists in SharePoint and delete it to enable replacement.

    This function implements the "delete-then-upload" strategy to ensure
    existing files are properly replaced with newer versions.

    Args:
        drive (DriveItem): The folder to check for existing file
        file_name (str): Name of the file to check (e.g., 'report.pdf')

    Returns:
        bool: True if an existing file was deleted, False if no file existed

    Example:
        was_deleted = check_and_delete_existing_file(folder, "data.xlsx")
        if was_deleted:
            print("Replacing existing file")
        else:
            print("Uploading new file")

    Note:
        This function is necessary because the Office365 library's upload_file()
        method doesn't overwrite existing files by default (known limitation).
        File names are sanitized for SharePoint compatibility before checking.
    """
    # Sanitize the file name to match what would be stored in SharePoint
    sanitized_name = sanitize_sharepoint_name(file_name, is_folder=False)

    try:
        # Attempt to retrieve file by sanitized name from SharePoint
        # get_by_path() navigates to the file, get() retrieves metadata
        # execute_query() sends the API request
        existing_file = drive.get_by_path(sanitized_name).get().execute_query()

        # Verify it's a file, not a folder with the same name
        # Files don't have a 'folder' attribute, folders do
        if not hasattr(existing_file, 'folder'):
            print(f"[!] Existing file found: {sanitized_name}")
            if sanitized_name != file_name:
                print(f"    (Original name: {file_name})")
            print(f"[×] Deleting existing file to prepare for replacement...")

            # Delete the file from SharePoint
            # delete_object() marks for deletion, execute_query() performs it
            existing_file.delete_object().execute_query()
            print(f"[✓] Existing file deleted successfully")

            # Brief pause to ensure SharePoint processes the deletion
            # Some SharePoint instances need this to avoid conflicts
            time.sleep(0.5)
            return True  # Signal that file was replaced
        else:
            # Edge case: A folder exists with the same name as our file
            print(f"[!] Found folder with same name as file: {file_name}")
            return False

    except Exception as e:
        # Exception usually means file doesn't exist (404 error)
        # This is expected for new files, so we return False
        # Other errors (network, permissions) will be caught later
        return False

def upload_file(drive, local_path, chunk_size, force_upload=False, desired_name=None):
    """
    Upload a file to SharePoint/OneDrive, intelligently skipping unchanged files.

    :param drive: The DriveItem representing the target folder
    :param local_path: Path to the local file to upload
    :param chunk_size: Size threshold for using resumable upload
    :param force_upload: If True, skip comparison and always upload with new hash
    :param desired_name: Optional desired filename in SharePoint (for temp file uploads)
    """
    # Use desired_name if provided (for HTML conversions), otherwise use actual filename
    file_name = desired_name if desired_name else os.path.basename(local_path)
    file_size = os.path.getsize(local_path)

    # Sanitize the file name for SharePoint compatibility
    sanitized_name = sanitize_sharepoint_name(file_name, is_folder=False)

    # Calculate file hash for later use
    local_hash = None

    # First, check if the file needs updating (unless forced)
    if not force_upload:
        needs_update, exists, remote_file, local_hash = check_file_needs_update(drive, local_path, file_name)

        # If file doesn't need updating, skip it
        if not needs_update:
            return  # File is identical, skip upload

        # If file exists but needs update, delete it first
        if exists and needs_update:
            print(f"[×] Deleting outdated file to prepare for update...")
            try:
                remote_file.delete_object().execute_query()
                print(f"[✓] Outdated file deleted successfully")
                time.sleep(0.5)  # Brief pause for SharePoint to process
                upload_stats['replaced_files'] += 1
            except Exception as e:
                print(f"[!] Warning: Could not delete existing file: {e}")

            print(f"[→] Uploading updated file: {sanitized_name}")
            if sanitized_name != file_name:
                print(f"    (Original name: {file_name})")
        else:
            # New file
            print(f"[→] Uploading new file: {sanitized_name}")
            if sanitized_name != file_name:
                print(f"    (Original name: {file_name})")
            upload_stats['new_files'] += 1
    else:
        # Force upload mode - always delete and reupload with new hash
        # Calculate hash now since we skipped check_file_needs_update
        local_hash = calculate_file_hash(local_path)
        if local_hash:
            print(f"[#] Calculated hash for force upload: {local_hash[:8]}...")

        file_was_deleted = check_and_delete_existing_file(drive, file_name)
        if file_was_deleted:
            print(f"[→] Force uploading replacement file: {sanitized_name}")
            upload_stats['replaced_files'] += 1
        else:
            print(f"[→] Force uploading new file: {sanitized_name}")
            upload_stats['new_files'] += 1

    try:
        # Special handling for files with periods in the name that might cause issues
        # If the file has multiple periods or is an AppxBundle, try direct upload first
        has_multiple_periods = file_name.count('.') > 1
        is_appx_file = file_name.lower().endswith(('.appxbundle', '.appx', '.msixbundle', '.msix'))

        # For problematic files, increase the chunk size threshold to 250MB
        # This forces them to use regular upload instead of resumable for files under 250MB
        effective_chunk_size = chunk_size
        if has_multiple_periods or is_appx_file:
            effective_chunk_size = 250 * 1024 * 1024  # 250MB
            print(f"[!] Special file detected, using direct upload for files under 250MB")

        # Create a temporary file with the sanitized name if needed
        temp_file_created = False
        upload_path = local_path

        # Create temp copy with the correct sanitized name for SharePoint
        if desired_name:
            # When we have a desired_name (e.g., for HTML conversions), always create temp with sanitized name
            # This ensures SharePoint gets the correct filename
            # Use a unique subdirectory to avoid conflicts between multiple files with same name
            temp_dir = tempfile.mkdtemp(prefix='sharepoint_upload_')
            temp_path = os.path.join(temp_dir, sanitized_name)
            shutil.copy2(local_path, temp_path)
            upload_path = temp_path
            temp_file_created = True
            # Store the temp dir for cleanup
            temp_dir_created = temp_dir
        elif sanitized_name != file_name:
            # For regular files, create temp copy only if sanitization changed the name
            temp_dir = tempfile.gettempdir()
            temp_path = os.path.join(temp_dir, sanitized_name)
            shutil.copy2(local_path, temp_path)
            upload_path = temp_path
            temp_file_created = True

        # Perform the upload based on file size
        if file_size < effective_chunk_size:
            remote_file = drive.upload_file(upload_path).execute_query()
            # Pass the desired name for display if it was provided
            display_name = desired_name if desired_name else os.path.basename(local_path)
            success_callback(remote_file, local_path, display_name=display_name)
        else:
            # resumable_upload handles sanitization internally
            # This is only used for very large files now
            resumable_upload(
                drive,
                local_path,  # Pass original path, function will sanitize
                file_size,
                chunk_size,
                max_chunk_retry=60,
                timeout_secs=10*60)

        # Clean up temporary file/directory if created
        if temp_file_created:
            if 'temp_dir_created' in locals() and os.path.exists(temp_dir_created):
                # Clean up the entire temp directory for HTML files
                shutil.rmtree(temp_dir_created)
                # Silent cleanup for normal operations
            elif os.path.exists(temp_path):
                # Clean up individual temp file for regular files
                os.remove(temp_path)
                # Silent cleanup for normal operations

        # Update upload byte counter after successful upload
        upload_stats['bytes_uploaded'] += file_size

        # Try to set the FileHash metadata if we have a hash using direct REST API
        if local_hash:
            try:
                print(f"[#] Setting FileHash metadata...")

                # Debug logging for FileHash setting
                debug_metadata = os.environ.get('DEBUG_METADATA', 'false').lower() == 'true'

                # First get the list item data to find the item ID
                list_item_data = get_sharepoint_list_item_by_filename(tenant_url, library_name, sanitized_name)

                if list_item_data and 'id' in list_item_data:
                    item_id = list_item_data['id']

                    if debug_metadata:
                        print(f"[DEBUG] Setting FileHash for {sanitized_name}")
                        print(f"[DEBUG] SharePoint list item ID: {item_id}")
                        print(f"[DEBUG] About to set FileHash to: {local_hash}")

                    # Update the FileHash field using REST API
                    success = update_sharepoint_list_item_field(
                        tenant_url,
                        library_name,
                        item_id,
                        'FileHash',
                        local_hash
                    )

                    if success:
                        print(f"[✓] FileHash metadata set: {local_hash[:8]}...")

                        # Debug logging to verify FileHash was set
                        if debug_metadata:
                            # Re-fetch to verify the FileHash was set correctly
                            verify_data = get_sharepoint_list_item_by_filename(tenant_url, library_name, sanitized_name)
                            if verify_data and 'fields' in verify_data:
                                verified_hash = verify_data['fields'].get('FileHash')
                                print(f"[DEBUG] FileHash verification after setting: {verified_hash}")
                                print(f"[DEBUG] FileHash matches expected: {verified_hash == local_hash}")
                            else:
                                print(f"[DEBUG] Unable to verify FileHash - could not retrieve updated item")
                    else:
                        print(f"[!] Failed to set FileHash metadata via REST API")
                else:
                    print(f"[!] Could not find list item for uploaded file to set hash metadata")

            except Exception as hash_error:
                print(f"[!] Could not set FileHash metadata via REST API: {str(hash_error)[:200]}")
                # Continue anyway - file is uploaded successfully

    except Exception as e:
        # IMPORTANT: Clean up temp file even on failure to prevent conflicts
        if temp_file_created:
            try:
                if 'temp_dir_created' in locals() and os.path.exists(temp_dir_created):
                    # Clean up the entire temp directory for HTML files
                    shutil.rmtree(temp_dir_created)
                    print(f"[!] Cleaned up temp directory after error: {temp_dir_created}")
                elif os.path.exists(temp_path):
                    # Clean up individual temp file for regular files
                    os.remove(temp_path)
                    print(f"[!] Cleaned up temp file after error: {temp_path}")
            except Exception as cleanup_error:
                print(f"[!] Warning: Could not delete temp file/dir: {cleanup_error}")

        upload_stats['failed_files'] += 1
        raise e

def upload_file_with_structure(root_drive, local_file_path, base_path=""):
    """
    Upload a file maintaining its directory structure

    :param root_drive: The root drive in SharePoint where files should be uploaded
    :param local_file_path: The local path of the file to upload
    :param base_path: The base path to strip from the file path (for relative paths)

    Compatible with Office365-REST-Python-Client v2.6.2
    """
    # Get the relative path of the file
    if base_path:
        rel_path = os.path.relpath(local_file_path, base_path)
    else:
        rel_path = local_file_path

    # Normalize path separators for cross-platform compatibility
    rel_path = rel_path.replace('\\', '/')

    # Sanitize the entire relative path for SharePoint compatibility
    # This ensures both folder and file names are properly sanitized
    sanitized_rel_path = sanitize_path_components(rel_path)

    # Get the directory path and file name from sanitized path
    dir_path = os.path.dirname(sanitized_rel_path)
    file_name = os.path.basename(sanitized_rel_path)

    # Log if path was sanitized
    if sanitized_rel_path != rel_path:
        print(f"[!] Path sanitized for SharePoint: {rel_path} -> {sanitized_rel_path}")

    # If there's a directory structure, create it in SharePoint
    # Note: ensure_folder_exists will sanitize folder names internally
    if dir_path and dir_path != "." and dir_path != "":
        target_folder = ensure_folder_exists(root_drive, dir_path)
    else:
        target_folder = root_drive
    
    # Upload the file to the target folder
    print(f"[→] Processing file: {local_file_path}")
    for i in range(max_retry):
        try:
            upload_file(target_folder, local_file_path, 4*1024*1024, force_upload)
            break
        except Exception as e:
            print(f"[Error] Upload failed: {e}, {type(e)}")
            if i == max_retry - 1:
                print(f"[Error] Failed to upload {local_file_path} after {max_retry} attempts")
                raise e
            else:
                print(f"[!] Retrying upload... ({i+1}/{max_retry})")
                time.sleep(2)

# ====================================================================
# BASE PATH CALCULATION - For maintaining folder structure
# ====================================================================
# We need to determine the "base" path to strip from file paths
# This preserves the relative folder structure when uploading

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

# ====================================================================
# SHAREPOINT CONNECTION TEST AND COLUMN SETUP
# ====================================================================
# Verify we can connect to SharePoint before processing files

print("[*] Connecting to SharePoint...")
try:
    # Execute the query to test connection and permissions
    # This also initializes the root_drive object for use
    root_drive.get().execute_query()
    print(f"[✓] Connected to SharePoint at: {upload_path}")

    # Check and create FileHash column if needed
    # Try to determine the document library name from the upload path
    # Default to "Documents" or "Shared Documents" if not specified
    library_name = "Documents"  # Default document library name
    if upload_path and "/" in upload_path:
        # If upload_path starts with a library name, use it
        path_parts = upload_path.split("/")
        if path_parts[0]:
            library_name = path_parts[0]

    # Attempt to create the FileHash column for hash-based comparison
    filehash_column_available = check_and_create_filehash_column(tenant_url, library_name)
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

# ====================================================================
# MAIN UPLOAD LOOP - Process each file
# ====================================================================
# Iterate through all discovered files and upload them to SharePoint

# Track converted files to avoid uploading .md files when .html versions exist
converted_md_files = set()

for f in local_files:
    # Safety check: Verify item is still a file (not deleted/moved)
    if os.path.isfile(f):
        # Check if this is a markdown file and conversion is enabled
        if f.lower().endswith('.md') and convert_md_to_html:
            print(f"[MD] Converting markdown file: {f}")

            try:
                # Read the markdown file
                with open(f, 'r', encoding='utf-8') as md_file:
                    md_content = md_file.read()

                # Convert to HTML
                html_content = convert_markdown_to_html(md_content, os.path.basename(f))

                # Create HTML file in temp directory to avoid permission issues
                # Use a simple prefix to avoid confusion - the actual filename will be set during upload
                temp_html_fd, html_path = tempfile.mkstemp(suffix='.html', prefix='converted_md_')

                try:
                    # Write HTML file to temp location
                    with os.fdopen(temp_html_fd, 'w', encoding='utf-8') as html_file:
                        html_file.write(html_content)
                except Exception as write_error:
                    os.close(temp_html_fd)  # Ensure file descriptor is closed
                    raise write_error

                # Check if HTML needs updating before upload
                # Create a synthetic path that matches the original .md location but with .html extension
                original_html_path = f.replace('.md', '.html')

                # Get the size and hash of the newly converted HTML file
                html_file_size = os.path.getsize(html_path)
                html_hash = calculate_file_hash(html_path)
                print(f"[HTML] Converted HTML size: {html_file_size:,} bytes")
                if html_hash:
                    print(f"[#] HTML hash: {html_hash[:8]}...")

                # We'll use a modified version of upload_file_with_structure that accepts
                # a separate actual file path and desired upload path
                # First, get the relative path structure from the original markdown
                if base_path:
                    rel_path = os.path.relpath(original_html_path, base_path)
                else:
                    rel_path = original_html_path

                # Normalize and sanitize the path
                rel_path = rel_path.replace('\\', '/')
                sanitized_rel_path = sanitize_path_components(rel_path)

                # Get the directory path for the HTML file (same as markdown)
                dir_path = os.path.dirname(sanitized_rel_path)

                # Create folder structure if needed
                if dir_path and dir_path != "." and dir_path != "":
                    target_folder = ensure_folder_exists(root_drive, dir_path)
                else:
                    target_folder = root_drive

                # Check if HTML file exists in SharePoint and compare hashes or sizes
                desired_html_filename = os.path.basename(original_html_path)
                html_needs_update = True  # Default to uploading
                hash_comparison_used = False

                try:
                    # Check if HTML already exists in SharePoint
                    children = target_folder.children.get().select(["name", "size", "file", "id"]).execute_query()
                    html_found = False
                    for child in children:
                        child_name = getattr(child, 'name', None)
                        if child_name and child_name == desired_html_filename:
                            html_found = True

                            # First try hash comparison if available using direct REST API
                            if html_hash and filehash_column_available:
                                try:
                                    # Debug logging for HTML FileHash retrieval
                                    debug_metadata = os.environ.get('DEBUG_METADATA', 'false').lower() == 'true'

                                    # Use direct Graph API REST calls to get SharePoint list item with custom columns
                                    list_item_data = get_sharepoint_list_item_by_filename(tenant_url, library_name, desired_html_filename)

                                    if list_item_data and 'fields' in list_item_data:
                                        fields = list_item_data['fields']

                                        if debug_metadata:
                                            print(f"[DEBUG] Retrieving FileHash for HTML {desired_html_filename}")
                                            print(f"[DEBUG] HTML REST API list item data: {type(list_item_data)}")
                                            print(f"[DEBUG] HTML fields data: {type(fields)}")
                                            print(f"[DEBUG] HTML available field properties: {list(fields.keys())}")
                                            print(f"[DEBUG] HTML FileHash in properties: {'FileHash' in fields}")
                                            if 'FileHash' in fields:
                                                print(f"[DEBUG] HTML FileHash value: {fields.get('FileHash')}")

                                        # Try to get FileHash from the fields
                                        remote_hash = fields.get('FileHash')

                                        if remote_hash:
                                            hash_comparison_used = True
                                            if remote_hash == html_hash:
                                                print(f"[=] HTML unchanged (hash match): {desired_html_filename}")
                                                html_needs_update = False
                                                upload_stats['skipped_files'] += 1
                                                upload_stats['bytes_skipped'] += html_file_size
                                            else:
                                                print(f"[*] HTML changed (hash mismatch): {desired_html_filename}")
                                                upload_stats['replaced_files'] += 1
                                            break
                                        elif debug_metadata:
                                            print(f"[DEBUG] HTML FileHash not found in list item fields")
                                    elif debug_metadata:
                                        print(f"[DEBUG] Could not retrieve HTML list item data for {desired_html_filename}")

                                except Exception as html_hash_error:
                                    # Hash comparison failed, fall back to size
                                    if debug_metadata:
                                        print(f"[DEBUG] HTML hash comparison failed: {str(html_hash_error)[:100]}")
                                    pass

                            # Fall back to size comparison if hash wasn't available
                            if not hash_comparison_used:
                                # Found existing HTML file - try multiple ways to get size
                                remote_size = None

                                # Try Graph API DriveItem properties
                                if hasattr(child, 'size') and child.size is not None:
                                    remote_size = child.size
                                # Try SharePoint File properties
                                elif hasattr(child, 'length') and child.length is not None:
                                    remote_size = child.length
                                # Try properties dictionary
                                elif hasattr(child, 'properties'):
                                    remote_size = child.properties.get('size') or child.properties.get('Size') or child.properties.get('length') or child.properties.get('Length')

                                if remote_size is not None:
                                    # Compare sizes - less reliable for HTML due to conversion variations
                                    if remote_size == html_file_size:
                                        print(f"[=] HTML unchanged (size: {html_file_size:,} bytes): {desired_html_filename}")
                                        html_needs_update = False
                                        upload_stats['skipped_files'] += 1
                                        upload_stats['bytes_skipped'] += html_file_size
                                    else:
                                        print(f"[*] HTML size changed (local: {html_file_size:,} vs remote: {remote_size:,}): {desired_html_filename}")
                                        upload_stats['replaced_files'] += 1
                                else:
                                    # Could not get size, assume needs update
                                    print(f"[!] Cannot determine remote HTML size, will upload: {desired_html_filename}")
                            break

                    if not html_found:
                        print(f"[+] New HTML file to upload: {desired_html_filename}")
                        upload_stats['new_files'] += 1
                except Exception as check_error:
                    print(f"[!] Could not check existing HTML, will upload: {check_error}")

                # Only upload if the HTML needs updating
                if html_needs_update:
                    print(f"[→] Processing file: {original_html_path} (from temp: {html_path})")
                    for i in range(max_retry):
                        try:
                            upload_file(target_folder, html_path, 4*1024*1024, force_upload, desired_name=desired_html_filename)
                            break
                        except Exception as e:
                            print(f"[Error] Upload failed: {e}, {type(e)}")
                            if i == max_retry - 1:
                                print(f"[Error] Failed to upload {original_html_path} after {max_retry} attempts")
                                raise e
                            else:
                                print(f"[!] Retrying upload... ({i+1}/{max_retry})")
                                time.sleep(2)
                else:
                    print(f"[✓] Skipping HTML upload - file is identical in SharePoint")

                # Clean up temporary HTML file (whether uploaded or skipped)
                if os.path.exists(html_path):
                    os.remove(html_path)

                # Mark this markdown file as converted
                converted_md_files.add(f)

            except Exception as e:
                print(f"[Error] Failed to convert markdown file {f}: {e}")
                print(f"[!] Uploading original markdown file instead")
                # Fall back to uploading the markdown file as-is
                upload_file_with_structure(root_drive, f, base_path)

        elif f.lower().endswith('.md') and not convert_md_to_html:
            # Markdown conversion is disabled, upload as-is
            print(f"[MD] Uploading markdown file as-is (conversion disabled): {f}")
            upload_file_with_structure(root_drive, f, base_path)

        elif f not in converted_md_files:
            # Regular file, not markdown or not converted
            upload_file_with_structure(root_drive, f, base_path)
    else:
        # File might have been deleted/moved since discovery
        print(f"[Warning] Skipping {f} as it's not a file")

# ====================================================================
# FINAL SUMMARY REPORT
# ====================================================================
# Display upload statistics to the user/CI system

# Create visual separator for better readability
print("\n" + "="*60)
print("[✓] SYNC PROCESS COMPLETED")
print("="*60)

# Calculate data sizes for display
def format_bytes(bytes):
    """Convert bytes to human-readable format"""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if bytes < 1024.0:
            return f"{bytes:.1f} {unit}"
        bytes /= 1024.0
    return f"{bytes:.1f} TB"

# Show detailed statistics
print(f"[STATS] Sync Statistics:")
print(f"   - New files uploaded:       {upload_stats['new_files']:>6}")
print(f"   - Files updated:            {upload_stats['replaced_files']:>6}")
print(f"   - Files skipped (unchanged):{upload_stats['skipped_files']:>6}")
print(f"   - Failed uploads:           {upload_stats['failed_files']:>6}")
print(f"   - Total files processed:    {len(local_files):>6}")
print(f"\n[DATA] Transfer Summary:")
print(f"   - Data uploaded:   {format_bytes(upload_stats['bytes_uploaded'])}")
print(f"   - Data skipped:    {format_bytes(upload_stats['bytes_skipped'])}")
print(f"   - Total savings:   {format_bytes(upload_stats['bytes_skipped'])} ({upload_stats['skipped_files']} files not re-uploaded)")

# Calculate efficiency percentage
total_processed = upload_stats['new_files'] + upload_stats['replaced_files'] + upload_stats['skipped_files']
if total_processed > 0:
    efficiency = (upload_stats['skipped_files'] / total_processed) * 100
    print(f"\n[EFFICIENCY] {efficiency:.1f}% of files were already up-to-date")

print("="*60)

# ====================================================================
# EXIT CODE HANDLING - For CI/CD integration
# ====================================================================
# Return appropriate exit code for GitHub Actions or other CI systems
# Exit code 0 = success, 1 = failure

if upload_stats['failed_files'] > 0:
    # Some files failed - signal error to CI system
    sys.exit(1)

# If we get here, all uploads succeeded (exit code 0 is implicit)
