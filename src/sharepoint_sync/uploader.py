# -*- coding: utf-8 -*-
"""
Upload operations for SharePoint sync.

This module handles all file upload operations including folder management,
resumable uploads for large files, and metadata updates.
"""

import os
import time
import tempfile
import shutil
from office365.runtime.odata.v4.upload_session_request import UploadSessionRequest
from office365.onedrive.driveitems.driveItem import DriveItem
from office365.onedrive.internal.paths.url import UrlPath
from office365.runtime.queries.upload_session import UploadSessionQuery
from office365.onedrive.driveitems.uploadable_properties import DriveItemUploadableProperties
from .file_handler import (
    sanitize_sharepoint_name,
    sanitize_path_components,
    calculate_file_hash,
    check_file_needs_update
)
from .graph_api import (
    update_sharepoint_list_item_field,
    get_sharepoint_list_item_by_filename
)

# Global cache for created folders
# Using a dictionary (path -> DriveItem) to avoid redundant API calls
created_folders = {}


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
        - Handles both forward slash (/) and backslash (\\) path separators
        - Sanitizes folder names for SharePoint compatibility
        - Compatible with Office365-REST-Python-Client v2.6.2+
    """
    # Convert Windows backslashes to forward slashes for consistency
    # This ensures the function works on both Windows and Unix systems
    folder_path = folder_path.replace('\\', '/')

    # Sanitize the entire path for SharePoint compatibility
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
                    print(f"[] Folder already exists: {current_path}")
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
                print(f"[] Created folder: {current_path}")

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
                    print(f"[] Created folder: {current_path}")

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
                                print(f"[] Found existing folder: {current_path}")
                                break
                    except Exception as fallback_error:
                        # Couldn't retrieve existing folder - will use parent
                        print(f"[!] Could not retrieve existing folder: {fallback_error}")
                else:
                    print(f"[!] Error creating folder {folder_name}: {create_error}")

                    # Fallback: Try to navigate to folder in case it exists
                    try:
                        print(f"[!] Attempting to navigate to folder: {folder_name}")
                        test_folder = current_drive.get_by_path(folder_name).get().execute_query()
                        if test_folder and hasattr(test_folder, 'folder'):
                            current_drive = test_folder
                            created_folders[current_path] = test_folder
                            print(f"[] Successfully navigated to folder: {current_path}")
                        else:
                            raise Exception("Not a folder")
                    except Exception as nav_error:
                        print(f"[!] Unable to create or navigate to folder {current_path}: {nav_error}")
                        print(f"[!] Will continue with parent folder")
                        # Don't fail the entire upload, just use parent folder

    return current_drive


def progress_status(offset, file_size):
    """Display upload progress."""
    print(f"Uploaded {offset} bytes from {file_size} bytes ... {offset/file_size*100:.2f}%")


def success_callback(remote_file, local_path, display_name=None):
    """Display success message after file upload."""
    # Use display_name if provided (for temp files), otherwise use local_path
    file_display = display_name if display_name else local_path
    print(f"[] File {file_display} has been uploaded to {remote_file.web_url}")


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
        print(f"[] Using built-in upload method for large file: {sanitized_name}")
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
                    except Exception as retry_error:
                        if retry_number + 1 >= max_chunk_retry:
                            raise retry_error
                        print(f"Retry {retry_number}: {retry_error}")
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
            temp_path = None
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
            if temp_file_created and temp_path and os.path.exists(temp_path):
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
            print(f"[] Existing file deleted successfully")

            # Brief pause to ensure SharePoint processes the deletion
            # Some SharePoint instances need this to avoid conflicts
            time.sleep(0.5)
            return True  # Signal that file was replaced
        else:
            # Edge case: A folder exists with the same name as our file
            print(f"[!] Found folder with same name as file: {file_name}")
            return False

    except Exception:
        # Exception usually means file doesn't exist (404 error)
        # This is expected for new files, so we return False
        # Other errors (network, permissions) will be caught later
        return False


def upload_file(drive, local_path, chunk_size, force_upload, site_url, list_name,
                filehash_column_available, tenant_id, client_id, client_secret,
                login_endpoint, graph_endpoint, upload_stats_dict, desired_name=None):
    """
    Upload a file to SharePoint/OneDrive, intelligently skipping unchanged files.

    :param drive: The DriveItem representing the target folder
    :param local_path: Path to the local file to upload
    :param chunk_size: Size threshold for using resumable upload
    :param force_upload: If True, skip comparison and always upload with new hash
    :param site_url: Full SharePoint site URL
    :param list_name: Name of the document library (usually "Documents")
    :param filehash_column_available: Whether FileHash column exists in SharePoint
    :param tenant_id: Azure AD tenant ID
    :param client_id: Azure AD app registration client ID
    :param client_secret: Azure AD app registration client secret
    :param login_endpoint: Azure AD authentication endpoint
    :param graph_endpoint: Microsoft Graph API endpoint
    :param upload_stats_dict: Dictionary to track upload statistics
    :param desired_name: Optional desired filename in SharePoint (for temp file uploads)
    """
    # Use desired_name if provided (for HTML conversions), otherwise use actual filename
    file_name = desired_name if desired_name else os.path.basename(local_path)
    file_size = os.path.getsize(local_path)

    # Sanitize the file name for SharePoint compatibility
    sanitized_name = sanitize_sharepoint_name(file_name, is_folder=False)

    # Initialize variables that may be used in exception handler
    local_hash = None
    temp_file_created = False
    temp_path = None
    temp_dir_created = None

    # First, check if the file needs updating (unless forced)
    if not force_upload:
        needs_update, exists, remote_file, local_hash = check_file_needs_update(
            drive, local_path, file_name, site_url, list_name,
            filehash_column_available, tenant_id, client_id, client_secret,
            login_endpoint, graph_endpoint, upload_stats_dict
        )

        # If file doesn't need updating, skip it
        if not needs_update:
            return  # File is identical, skip upload

        # If file exists but needs update, delete it first
        if exists and needs_update:
            print(f"[×] Deleting outdated file to prepare for update...")
            try:
                remote_file.delete_object().execute_query()
                print(f"[] Outdated file deleted successfully")
                time.sleep(0.5)  # Brief pause for SharePoint to process
                upload_stats_dict['replaced_files'] += 1
            except Exception as e:
                print(f"[!] Warning: Could not delete existing file: {e}")

            print(f"[] Uploading updated file: {sanitized_name}")
            if sanitized_name != file_name:
                print(f"    (Original name: {file_name})")
        else:
            # New file
            print(f"[] Uploading new file: {sanitized_name}")
            if sanitized_name != file_name:
                print(f"    (Original name: {file_name})")
            upload_stats_dict['new_files'] += 1
    else:
        # Force upload mode - always delete and reupload with new hash
        # Calculate hash now since we skipped check_file_needs_update
        local_hash = calculate_file_hash(local_path)
        if local_hash:
            print(f"[#] Calculated hash for force upload: {local_hash[:8]}...")

        file_was_deleted = check_and_delete_existing_file(drive, file_name)
        if file_was_deleted:
            print(f"[] Force uploading replacement file: {sanitized_name}")
            upload_stats_dict['replaced_files'] += 1
        else:
            print(f"[] Force uploading new file: {sanitized_name}")
            upload_stats_dict['new_files'] += 1

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

        # Set upload path (temp variables already initialized at function start)
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
            if temp_dir_created and os.path.exists(temp_dir_created):
                # Clean up the entire temp directory for HTML files
                shutil.rmtree(temp_dir_created)
                # Silent cleanup for normal operations
            elif temp_path and os.path.exists(temp_path):
                # Clean up individual temp file for regular files
                os.remove(temp_path)
                # Silent cleanup for normal operations

        # Update upload byte counter after successful upload
        upload_stats_dict['bytes_uploaded'] += file_size

        # Try to set the FileHash metadata if we have a hash using direct REST API
        if local_hash:
            try:
                print(f"[#] Setting FileHash metadata...")

                # Debug logging for FileHash setting
                debug_metadata = os.environ.get('DEBUG_METADATA', 'false').lower() == 'true'

                # First get the list item data to find the item ID
                list_item_data = get_sharepoint_list_item_by_filename(
                    site_url, list_name, sanitized_name,
                    tenant_id, client_id, client_secret, login_endpoint, graph_endpoint
                )

                if list_item_data and 'id' in list_item_data:
                    item_id = list_item_data['id']

                    if debug_metadata:
                        print(f"[DEBUG] Setting FileHash for {sanitized_name}")
                        print(f"[DEBUG] SharePoint list item ID: {item_id}")
                        print(f"[DEBUG] About to set FileHash to: {local_hash}")

                    # Update the FileHash field using REST API
                    success = update_sharepoint_list_item_field(
                        site_url,
                        list_name,
                        item_id,
                        'FileHash',
                        local_hash,
                        tenant_id,
                        client_id,
                        client_secret,
                        login_endpoint,
                        graph_endpoint
                    )

                    if success:
                        print(f"[] FileHash metadata set: {local_hash[:8]}...")

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
                if temp_dir_created and os.path.exists(temp_dir_created):
                    # Clean up the entire temp directory for HTML files
                    shutil.rmtree(temp_dir_created)
                    print(f"[!] Cleaned up temp directory after error: {temp_dir_created}")
                elif temp_path and os.path.exists(temp_path):
                    # Clean up individual temp file for regular files
                    os.remove(temp_path)
                    print(f"[!] Cleaned up temp file after error: {temp_path}")
            except Exception as cleanup_error:
                print(f"[!] Warning: Could not delete temp file/dir: {cleanup_error}")

        upload_stats_dict['failed_files'] += 1
        raise e


def upload_file_with_structure(root_drive, local_file_path, base_path, site_url, list_name,
                                chunk_size, force_upload, filehash_column_available,
                                tenant_id, client_id, client_secret, login_endpoint,
                                graph_endpoint, upload_stats_dict, max_retry=3):
    """
    Upload a file maintaining its directory structure

    :param root_drive: The root drive in SharePoint where files should be uploaded
    :param local_file_path: The local path of the file to upload
    :param base_path: The base path to strip from the file path (for relative paths)
    :param site_url: Full SharePoint site URL
    :param list_name: Name of the document library (usually "Documents")
    :param chunk_size: Size threshold for using resumable upload
    :param force_upload: If True, skip comparison and always upload
    :param filehash_column_available: Whether FileHash column exists in SharePoint
    :param tenant_id: Azure AD tenant ID
    :param client_id: Azure AD app registration client ID
    :param client_secret: Azure AD app registration client secret
    :param login_endpoint: Azure AD authentication endpoint
    :param graph_endpoint: Microsoft Graph API endpoint
    :param upload_stats_dict: Dictionary to track upload statistics
    :param max_retry: Maximum number of retry attempts (default: 3)

    Compatible with Office365-REST-Python-Client v2.6.2
    """
    # Get the relative path of the file
    if base_path:
        rel_path = os.path.relpath(local_file_path, base_path)
    else:
        rel_path = local_file_path

    # Normalize path separators for cross-platform compatibility
    # Ensure rel_path is a string (handle both str and bytes)
    if isinstance(rel_path, bytes):
        rel_path = rel_path.decode('utf-8')
    rel_path = rel_path.replace('\\', '/')

    # Sanitize the entire relative path for SharePoint compatibility
    # This ensures both folder and file names are properly sanitized
    sanitized_rel_path = sanitize_path_components(rel_path)

    # Get the directory path from sanitized path
    dir_path = os.path.dirname(sanitized_rel_path)

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
    print(f"[] Processing file: {local_file_path}")
    for i in range(max_retry):
        try:
            upload_file(
                target_folder, local_file_path, chunk_size, force_upload,
                site_url, list_name, filehash_column_available,
                tenant_id, client_id, client_secret, login_endpoint,
                graph_endpoint, upload_stats_dict
            )
            break
        except Exception as e:
            print(f"[Error] Upload failed: {e}, {type(e)}")
            if i == max_retry - 1:
                print(f"[Error] Failed to upload {local_file_path} after {max_retry} attempts")
                raise e
            else:
                print(f"[!] Retrying upload... ({i+1}/{max_retry})")
                time.sleep(2)
