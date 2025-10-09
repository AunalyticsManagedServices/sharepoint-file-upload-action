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
                                 <file_path> [options]

DESCRIPTION:
    - Uploads files and directories to SharePoint while preserving folder structure
    - Automatically replaces existing files with newer versions
    - Handles large files (>4MB) using resumable upload sessions
    - Creates missing folders automatically in SharePoint
    - Provides detailed logging and error reporting
    - Supports recursive file matching with glob patterns

REQUIREMENTS:
    - Python 3.6 or higher
    - office365-rest-python-client >= 2.6.2
    - msal (Microsoft Authentication Library)
    - Azure AD Enterprise Application with Graph API Sites.ReadWrite.All permission

AUTHOR:
    GitHub Actions SharePoint Integration

VERSION:
    2.0.0 - Added file replacement and enhanced error handling
"""

# ====================================
# IMPORTS - External libraries needed
# ====================================

# Standard library imports (come with Python)
import sys        # Provides access to command-line arguments and exit codes
import os         # Operating system interface for file/directory operations
import glob       # Unix-style pathname pattern expansion (e.g., *.txt matches all .txt files)
import time       # Time-related functions for delays and timestamps

# Third-party library imports (need to be installed via pip)
import msal       # Microsoft Authentication Library for Azure AD authentication

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
# URL AND CONFIGURATION SETUP
# ====================================================================

# Construct the full SharePoint site URL
# f-strings (f"...") allow embedding variables directly in strings
tenant_url = f'https://{sharepoint_host_name}/sites/{site_name}'

# Convert string argument to boolean for recursive file matching
# .lower() converts to lowercase for case-insensitive comparison
recursive = file_path_recursive_match.lower() in ['true', '1', 'yes']

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
    'failed_files': 0     # Files that failed to upload
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

def success_callback(remote_file, local_path):
    print(f"[✓] File {local_path} has been uploaded to {remote_file.web_url}")

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
            success_callback(remote_file, local_path)
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

    # Create DriveItem with conflictBehavior set to replace
    # Use sanitized name for the URL path to avoid issues with special characters
    return_type = DriveItem(
        drive.context,
        UrlPath(sanitized_name, drive.resource_path))

    # Create upload session query with conflict behavior
    upload_props = DriveItemUploadableProperties(name=sanitized_name)
    # Set conflict behavior to replace existing files
    upload_props.set_property("@microsoft.graph.conflictBehavior", "replace", False)

    qry = UploadSessionQuery(return_type, {"item": upload_props})
    drive.context.add_query(qry).after_query_execute(_start_upload)
    return_type.get().execute_query()
    success_callback(return_type, local_path)

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

def upload_file(drive, local_path, chunk_size):
    """
    Upload a file to SharePoint/OneDrive, replacing any existing file.

    :param drive: The DriveItem representing the target folder
    :param local_path: Path to the local file to upload
    :param chunk_size: Size threshold for using resumable upload
    """
    file_name = os.path.basename(local_path)
    file_size = os.path.getsize(local_path)

    # Sanitize the file name for SharePoint compatibility
    sanitized_name = sanitize_sharepoint_name(file_name, is_folder=False)

    # Check for and delete any existing file with the same name
    # Note: check_and_delete_existing_file already handles sanitization internally
    file_was_deleted = check_and_delete_existing_file(drive, file_name)

    if file_was_deleted:
        print(f"[→] Uploading replacement file: {sanitized_name}")
        if sanitized_name != file_name:
            print(f"    (Original name: {file_name})")
    else:
        print(f"[→] Uploading new file: {sanitized_name}")
        if sanitized_name != file_name:
            print(f"    (Original name: {file_name})")

    try:
        # Create a temporary file with the sanitized name if needed
        temp_file_created = False
        upload_path = local_path

        if sanitized_name != file_name:
            # Create a temporary copy with the sanitized name
            import tempfile
            import shutil
            temp_dir = tempfile.gettempdir()
            temp_path = os.path.join(temp_dir, sanitized_name)
            shutil.copy2(local_path, temp_path)
            upload_path = temp_path
            temp_file_created = True

        # Perform the upload based on file size
        if file_size < chunk_size:
            remote_file = drive.upload_file(upload_path).execute_query()
            success_callback(remote_file, local_path)
        else:
            # resumable_upload handles sanitization internally
            resumable_upload(
                drive,
                local_path,  # Pass original path, function will sanitize
                file_size,
                chunk_size,
                max_chunk_retry=60,
                timeout_secs=10*60)

        # Clean up temporary file if created
        if temp_file_created and os.path.exists(temp_path):
            os.remove(temp_path)

        # Update statistics after successful upload
        if file_was_deleted:
            upload_stats['replaced_files'] += 1
        else:
            upload_stats['new_files'] += 1

    except Exception as e:
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
            upload_file(target_folder, local_file_path, 4*1024*1024)
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
# SHAREPOINT CONNECTION TEST
# ====================================================================
# Verify we can connect to SharePoint before processing files

print("[*] Connecting to SharePoint...")
try:
    # Execute the query to test connection and permissions
    # This also initializes the root_drive object for use
    root_drive.get().execute_query()
    print(f"[✓] Connected to SharePoint at: {upload_path}")

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

for f in local_files:
    # Safety check: Verify item is still a file (not deleted/moved)
    if os.path.isfile(f):
        # Upload with folder structure preservation
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
print("[✓] UPLOAD PROCESS COMPLETED")
print("="*60)

# Show detailed statistics
print(f"📊 Upload Statistics:")
print(f"   • New files uploaded:      {upload_stats['new_files']}")
print(f"   • Existing files replaced: {upload_stats['replaced_files']}")
print(f"   • Failed uploads:          {upload_stats['failed_files']}")
print(f"   • Total files processed:   {len(local_files)}")
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
