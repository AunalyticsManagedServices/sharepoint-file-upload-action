# -*- coding: utf-8 -*-
"""
Upload operations for SharePoint sync.

This module handles all file upload operations including folder management,
resumable uploads for large files, and metadata updates.

All operations use direct Graph REST API calls.
"""

import os
import time
from .file_handler import (
    sanitize_sharepoint_name,
    sanitize_path_components,
    calculate_file_hash,
    check_file_needs_update
)
from .graph_api import (
    update_sharepoint_list_item_field,
    create_folder_graph,
    list_folder_children_graph,
    upload_small_file_graph,
    create_upload_session_graph,
    upload_file_chunk_graph
)
from .utils import is_debug_enabled

# Global cache for created folders
# Using a dictionary (path -> folder_item_dict) to avoid redundant API calls
# Structure: {path: {'id': item_id, 'name': folder_name, ...}}
created_folders = {}


def ensure_folder_exists(site_id, drive_id, parent_item_id, folder_path,
                        tenant_id, client_id, client_secret, login_endpoint, graph_endpoint,
                        folder_cache=None):
    """
    Recursively create folder structure in SharePoint if it doesn't exist using Graph API.

    This function handles nested folder creation, ensuring the entire path
    exists before uploading files. It uses caching to avoid redundant API calls.

    Args:
        site_id (str): SharePoint site ID
        drive_id (str): SharePoint drive ID
        parent_item_id (str): Parent folder item ID where structure should be created
        folder_path (str): Path to create (e.g., 'folder1/folder2/folder3')
        tenant_id (str): Azure AD tenant ID
        client_id (str): Azure AD application client ID
        client_secret (str): Azure AD application client secret
        login_endpoint (str): Azure AD login endpoint
        graph_endpoint (str): Microsoft Graph API endpoint
        folder_cache (dict): Optional folder cache from build_sharepoint_cache()
            Keys are folder paths, values are dicts with 'item_id' and 'name'

    Returns:
        str: The item ID of the final folder in the path, ready to receive files

    Raises:
        Exception: If folder creation fails after all retry attempts

    Example:
        target_id = ensure_folder_exists(site_id, drive_id, root_id, "2024/Reports/January", ..., folder_cache=cache['folders'])
        # Now upload file to the January folder using target_id

    Note:
        - Caches created folders to minimize API calls
        - Uses folder_cache (from SharePoint metadata cache) to skip folder existence checks
        - Handles both forward slash (/) and backslash (\\) path separators
        - Sanitizes folder names for SharePoint compatibility
        - Uses direct Graph REST API calls
    """
    # Convert Windows backslashes to forward slashes for consistency
    folder_path = folder_path.replace('\\', '/')

    # Sanitize the entire path for SharePoint compatibility
    folder_path = sanitize_path_components(folder_path)

    # Check cache first to avoid unnecessary API calls
    if folder_path in created_folders:
        return created_folders[folder_path]['id']

    # Split path into individual folder names
    path_parts = [part for part in folder_path.split('/') if part]

    # If no folders to create, return the parent
    if not path_parts:
        return parent_item_id

    # Start from the parent folder
    current_item_id = parent_item_id
    current_path = ""  # Track the path we've built so far

    # Process each folder in the path
    for folder_name in path_parts:
        # Build cumulative path as we go deeper
        current_path = f"{current_path}/{folder_name}" if current_path else folder_name

        # Skip if we've already processed this folder path
        if current_path in created_folders:
            current_item_id = created_folders[current_path]['id']
            continue

        # ============================================================
        # STEP 1: Check if folder already exists in SharePoint
        # ============================================================
        folder_found = False

        # Check folder cache first (fastest - 0 API calls)
        if folder_cache and current_path in folder_cache:
            folder_item = {
                'id': folder_cache[current_path]['item_id'],
                'name': folder_cache[current_path]['name']
            }
            created_folders[current_path] = folder_item
            current_item_id = folder_item['id']
            if is_debug_enabled():
                print(f"[CACHE HIT] Folder found in cache: {current_path}")
            folder_found = True

        # Fall back to API query if not in cache
        if not folder_found:
            try:
                if is_debug_enabled():
                    print(f"[?] Checking if folder exists: {current_path}")

                # Get all items in current folder using Graph API
                children = list_folder_children_graph(
                    site_id, drive_id, current_item_id,
                    tenant_id, client_id, client_secret, login_endpoint, graph_endpoint,
                    folder_path=current_path
                )

                if children is not None:
                    # Iterate through children to find matching folder
                    for child in children:
                        # Check if this is a folder with matching name
                        if child.get('name') == folder_name and 'folder' in child:
                            # Folder found! Cache it
                            folder_item = {
                                'id': child.get('id'),
                                'name': child.get('name')
                            }
                            created_folders[current_path] = folder_item
                            current_item_id = folder_item['id']
                            if is_debug_enabled():
                                print(f"[✓] Folder already exists: {current_path}")
                            folder_found = True
                            break

            except Exception as e:
                # API call failed - assume folder doesn't exist
                print(f"[!] Error checking folder existence: {e}")
                folder_found = False

        # ============================================================
        # STEP 2: Create folder if it doesn't exist
        # ============================================================
        if not folder_found:
            try:
                if is_debug_enabled():
                    print(f"[+] Creating folder: {folder_name}")

                # Create folder using Graph API
                created_folder = create_folder_graph(
                    site_id, drive_id, current_item_id, folder_name,
                    tenant_id, client_id, client_secret, login_endpoint, graph_endpoint
                )

                if created_folder:
                    folder_item = {
                        'id': created_folder.get('id'),
                        'name': created_folder.get('name')
                    }
                    created_folders[current_path] = folder_item
                    current_item_id = folder_item['id']
                    if is_debug_enabled():
                        print(f"[✓] Created folder: {current_path}")
                else:
                    raise Exception("Failed to create folder")

            except Exception as create_error:
                error_msg = str(create_error)

                # Check if folder already exists (common race condition)
                if "nameAlreadyExists" in error_msg or "already exists" in error_msg.lower():
                    if is_debug_enabled():
                        print(f"[!] Folder already exists (race condition): {folder_name}")
                    try:
                        # Try to get the existing folder
                        children = list_folder_children_graph(
                            site_id, drive_id, current_item_id,
                            tenant_id, client_id, client_secret, login_endpoint, graph_endpoint,
                            folder_path=current_path
                        )
                        if children:
                            for child in children:
                                if child.get('name') == folder_name and 'folder' in child:
                                    folder_item = {
                                        'id': child.get('id'),
                                        'name': child.get('name')
                                    }
                                    created_folders[current_path] = folder_item
                                    current_item_id = folder_item['id']
                                    if is_debug_enabled():
                                        print(f"[✓] Found existing folder: {current_path}")
                                    break
                    except Exception as fallback_error:
                        print(f"[!] Could not retrieve existing folder: {fallback_error}")
                else:
                    print(f"[!] Error creating folder {folder_name}: {create_error}")
                    print(f"[!] Will continue with parent folder")

    return current_item_id


def progress_status(offset, file_size):
    """Display upload progress."""
    if is_debug_enabled():
        print(f"Uploaded {offset} bytes from {file_size} bytes ... {offset/file_size*100:.2f}%")


def success_callback(remote_file, local_path, display_name=None, is_update=False):
    """
    Display success message after file upload.

    Args:
        remote_file: The uploaded file object
        local_path: Path to the local file
        display_name: Display name for temp files
        is_update: True if file was updated, False if newly processed
    """
    # Use display_name if provided (for temp files), otherwise use local_path
    file_display = display_name if display_name else os.path.basename(local_path)

    # Always show simple status message
    if is_update:
        print(f"File Updated: {file_display}")
    else:
        print(f"File Processed: {file_display}")

    # Show detailed URL only in DEBUG mode
    if is_debug_enabled():
        print(f"  → Uploaded to: {remote_file.web_url}")


def resumable_upload(site_id, drive_id, parent_item_id, local_path, filename, file_size, chunk_size,
                    tenant_id, client_id, client_secret, login_endpoint, graph_endpoint, is_update=False):
    """
    Upload large files using resumable upload sessions via Graph API.

    Args:
        site_id (str): SharePoint site ID
        drive_id (str): SharePoint drive ID
        parent_item_id (str): Parent folder item ID
        local_path (str): Path to the local file to upload
        filename (str): Desired filename in SharePoint
        file_size (int): Size of the file in bytes
        chunk_size (int): Size of each chunk to upload (must be multiple of 320 KiB)
        tenant_id (str): Azure AD tenant ID
        client_id (str): Azure AD application client ID
        client_secret (str): Azure AD application client secret
        login_endpoint (str): Azure AD login endpoint
        graph_endpoint (str): Microsoft Graph API endpoint
        is_update (bool): True if file is being updated, False if new

    Returns:
        dict: Uploaded file metadata from Graph API

    Note:
        - Uses Graph API createUploadSession endpoint
        - Supports files up to 250 GB
        - Chunk sizes must be multiples of 320 KiB (327,680 bytes)
        - Maximum 60 MiB per chunk
        - is_update parameter used for logging (shows "Updating" vs "Uploading")
    """
    sanitized_name = sanitize_sharepoint_name(filename, is_folder=False)

    if is_debug_enabled():
        action = "Updating" if is_update else "Uploading"
        print(f"[→] {action} large file with resumable upload: {sanitized_name} ({file_size:,} bytes)")
        if sanitized_name != filename:
            print(f"    (Original name: {filename})")

    try:
        # Step 1: Create upload session
        session = create_upload_session_graph(
            site_id, drive_id, parent_item_id, sanitized_name,
            tenant_id, client_id, client_secret, login_endpoint, graph_endpoint
        )

        if not session or 'uploadUrl' not in session:
            raise Exception("Failed to create upload session")

        upload_url = session['uploadUrl']

        # Step 2: Upload file in chunks
        #Ensure chunk size is multiple of 320 KiB (Graph API requirement)
        CHUNK_ALIGNMENT = 327680  # 320 KiB
        if chunk_size % CHUNK_ALIGNMENT != 0:
            chunk_size = ((chunk_size // CHUNK_ALIGNMENT) + 1) * CHUNK_ALIGNMENT

        # Cap at 60 MiB per Microsoft's recommendation
        MAX_CHUNK_SIZE = 60 * 1024 * 1024
        if chunk_size > MAX_CHUNK_SIZE:
            chunk_size = MAX_CHUNK_SIZE

        if is_debug_enabled():
            print(f"[DEBUG] Upload session created. Chunk size: {chunk_size:,} bytes")

        with open(local_path, 'rb') as f:
            offset = 0
            while offset < file_size:
                # Read next chunk
                f.seek(offset)
                chunk_data = f.read(chunk_size)
                chunk_end = offset + len(chunk_data) - 1

                # Upload chunk
                result = upload_file_chunk_graph(
                    upload_url, chunk_data, offset, chunk_end, file_size
                )

                if result is None:
                    raise Exception(f"Failed to upload chunk at offset {offset}")

                # Update progress
                progress_status(offset + len(chunk_data), file_size)

                offset += len(chunk_data)

                # Check if upload is complete
                if 'id' in result:
                    # Upload complete! File metadata returned
                    if is_debug_enabled():
                        print(f"[✓] Large file upload complete: {sanitized_name}")
                    return result

        # If we get here, all chunks uploaded successfully
        if is_debug_enabled():
            print(f"[✓] All chunks uploaded successfully for: {sanitized_name}")

        return {'name': sanitized_name, 'size': file_size}

    except Exception as e:
        print(f"[!] Resumable upload failed for {sanitized_name}: {e}")
        raise


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
            if is_debug_enabled():
                print(f"[!] Existing file found: {sanitized_name}")
            if sanitized_name != file_name:
                if is_debug_enabled():
                    print(f"    (Original name: {file_name})")
            if is_debug_enabled():
                print(f"[×] Deleting existing file to prepare for replacement...")

            # Delete the file from SharePoint
            # delete_object() marks for deletion, execute_query() performs it
            existing_file.delete_object().execute_query()
            if is_debug_enabled():
                print(f"[✓] Existing file deleted successfully")

            # Brief pause to ensure SharePoint processes the deletion
            # Some SharePoint instances need this to avoid conflicts
            time.sleep(0.5)
            return True  # Signal that file was replaced
        else:
            # Edge case: A folder exists with the same name as our file
            if is_debug_enabled():
                print(f"[!] Found folder with same name as file: {file_name}")
            return False

    except Exception:  # noqa: S110 - Broad exception acceptable here
        # Exception usually means file doesn't exist (404 error)
        # This is expected for new files, so we return False
        # Other errors (network, permissions) will be caught later during upload
        return False


def upload_file(site_id, drive_id, parent_item_id, local_path, chunk_size, force_upload, site_url, list_name,
                filehash_column_available, tenant_id, client_id, client_secret,
                login_endpoint, graph_endpoint, upload_stats_dict, desired_name=None,
                metadata_queue=None, pre_calculated_hash=None, display_path=None, sharepoint_cache=None):
    """
    Upload a file to SharePoint using Graph API, intelligently skipping unchanged files.

    Args:
        site_id (str): SharePoint site ID
        drive_id (str): SharePoint drive ID
        parent_item_id (str): Parent folder item ID
        local_path (str): Path to the local file to upload
        chunk_size (int): Size threshold for using resumable upload (250 MB per Graph API limit)
        force_upload (bool): If True, skip comparison and always upload with new hash
        site_url (str): Full SharePoint site URL
        list_name (str): Name of the document library (usually "Documents")
        filehash_column_available (bool): Whether FileHash column exists in SharePoint
        tenant_id (str): Azure AD tenant ID
        client_id (str): Azure AD app registration client ID
        client_secret (str): Azure AD app registration client secret
        login_endpoint (str): Azure AD authentication endpoint
        graph_endpoint (str): Microsoft Graph API endpoint
        upload_stats_dict (dict): Dictionary to track upload statistics
        desired_name (str): Optional desired filename in SharePoint (for temp file uploads)
        metadata_queue: Optional BatchQueue for batching metadata updates (parallel mode)
        pre_calculated_hash (str): Optional pre-calculated hash to use (for converted markdown using source .md hash)
        display_path (str): Optional relative path for display in debug output (e.g., 'docs/api/README.html')
        sharepoint_cache (dict): Optional pre-built cache of SharePoint file metadata (eliminates API calls for comparison)
    """
    # Use desired_name if provided (for HTML conversions), otherwise use actual filename
    file_name = desired_name if desired_name else os.path.basename(local_path)
    file_size = os.path.getsize(local_path)

    # Sanitize the file name for SharePoint compatibility
    sanitized_name = sanitize_sharepoint_name(file_name, is_folder=False)

    # Initialize variables
    local_hash = None
    is_file_update = False  # Track if this is an update vs new file

    # First, check if the file needs updating (unless forced)
    if not force_upload:
        # Note: check_file_needs_update now uses Graph API internally or cache lookup
        # Pass pre_calculated_hash if provided (for converted markdown using source .md hash)
        # Pass site_id, drive_id, parent_item_id for path-based queries (fixes duplicate filename bug)
        # Pass sharepoint_cache for instant lookups (eliminates API calls)
        needs_update, exists, remote_file, local_hash = check_file_needs_update(
            local_path, file_name, site_url, list_name,
            filehash_column_available, tenant_id, client_id, client_secret,
            login_endpoint, graph_endpoint, upload_stats_dict, pre_calculated_hash, display_path,
            site_id, drive_id, parent_item_id, sharepoint_cache
        )

        # If file doesn't need updating, skip it
        if not needs_update:
            return  # File is identical, skip upload

        # If file exists but needs update, we'll just replace it (Graph API handles conflict)
        if exists and needs_update:
            is_file_update = True
            if is_debug_enabled():
                display_name = display_path if display_path else file_name
                print(f"[→] Uploading updated file: {display_name}")
                if sanitized_name != file_name:
                    print(f"    (Original name: {file_name})")
            upload_stats_dict['replaced_files'] += 1
        else:
            # New file
            if is_debug_enabled():
                print(f"[→] Uploading new file: {sanitized_name}")
                if sanitized_name != file_name:
                    print(f"    (Original name: {file_name})")
            upload_stats_dict['new_files'] += 1
    else:
        # Force upload mode - always upload with new hash
        # Use pre_calculated_hash if provided, otherwise calculate from file
        if pre_calculated_hash:
            local_hash = pre_calculated_hash
            if is_debug_enabled():
                print(f"[#] Using pre-calculated hash for force upload: {local_hash[:8]}...")
        else:
            local_hash = calculate_file_hash(local_path)
            if local_hash and is_debug_enabled():
                print(f"[#] Calculated hash for force upload: {local_hash[:8]}...")

        # Check if file exists by listing children
        try:
            children = list_folder_children_graph(
                site_id, drive_id, parent_item_id,
                tenant_id, client_id, client_secret, login_endpoint, graph_endpoint
            )
            file_exists = False
            if children:
                for child in children:
                    if child.get('name') == sanitized_name and 'file' in child:
                        file_exists = True
                        break

            if file_exists:
                is_file_update = True
                if is_debug_enabled():
                    print(f"[→] Force uploading replacement file: {sanitized_name}")
                upload_stats_dict['replaced_files'] += 1
            else:
                if is_debug_enabled():
                    print(f"[→] Force uploading new file: {sanitized_name}")
                upload_stats_dict['new_files'] += 1
        except Exception as check_error:
            if is_debug_enabled():
                print(f"[!] Could not check file existence: {check_error}")
            # Assume new file
            upload_stats_dict['new_files'] += 1

    try:
        # Perform the upload based on file size
        # Graph API supports up to 250 MB for simple upload, use sessions for larger
        GRAPH_SMALL_FILE_LIMIT = 250 * 1024 * 1024  # 250 MB

        uploaded_item = None

        if file_size < GRAPH_SMALL_FILE_LIMIT:
            # Small file - use simple upload
            if is_debug_enabled():
                action = "Updating" if is_file_update else "Uploading"
                display_name = display_path if display_path else file_name
                print(f"[→] {action} file with simple upload: {display_name} ({file_size:,} bytes)")

            # Read file content
            with open(local_path, 'rb') as f:
                file_content = f.read()

            # Upload using Graph API
            uploaded_item = upload_small_file_graph(
                site_id, drive_id, parent_item_id, sanitized_name, file_content,
                tenant_id, client_id, client_secret, login_endpoint, graph_endpoint
            )

            # Verify upload succeeded
            if uploaded_item:
                if is_debug_enabled():
                    result_action = "updated" if is_file_update else "uploaded"
                    print(f"[✓] File {result_action} successfully: {sanitized_name}")
            else:
                raise Exception("Upload failed - no item returned")

        else:
            # Large file - use resumable upload
            uploaded_item = resumable_upload(
                site_id, drive_id, parent_item_id,
                local_path,  # Pass original path
                sanitized_name,  # Pass sanitized filename
                file_size,
                chunk_size,
                tenant_id, client_id, client_secret,
                login_endpoint, graph_endpoint,
                is_update=is_file_update
            )

        # Update upload byte counter after successful upload
        upload_stats_dict['bytes_uploaded'] += file_size

        # Use pre_calculated_hash if provided, otherwise use local_hash from check or force mode
        hash_to_save = pre_calculated_hash if pre_calculated_hash else local_hash

        # Try to set the FileHash metadata if we have a hash
        if hash_to_save and uploaded_item:
            try:
                # Get list item ID from the uploaded drive item
                # We know exactly where we uploaded (parent_item_id + sanitized_name),
                # so we can query the Graph API by path to get the listItem ID.
                item_id = None

                # Primary method: Query by path (most direct since we know the location)
                if is_debug_enabled():
                    print(f"[DEBUG] Fetching list item ID by path: parent={parent_item_id}, file={sanitized_name}")

                try:
                    from .graph_api import get_drive_item_by_path_with_list_item
                    item_with_list = get_drive_item_by_path_with_list_item(
                        site_id, drive_id, parent_item_id, sanitized_name,
                        tenant_id, client_id, client_secret, login_endpoint, graph_endpoint
                    )
                    if item_with_list and 'listItem' in item_with_list and 'id' in item_with_list['listItem']:
                        item_id = item_with_list['listItem']['id']
                        if is_debug_enabled():
                            print(f"[DEBUG] Got list item ID by path: {item_id}")
                except Exception as fetch_error:
                    if is_debug_enabled():
                        print(f"[DEBUG] Failed to fetch by path: {str(fetch_error)[:200]}")

                    # Fallback: If upload response has listItem (unlikely but check)
                    if 'listItem' in uploaded_item and 'id' in uploaded_item['listItem']:
                        item_id = uploaded_item['listItem']['id']
                        if is_debug_enabled():
                            print(f"[DEBUG] Got list item ID from upload response: {item_id}")
                    # Fallback: Try fetching by drive item ID
                    elif 'id' in uploaded_item:
                        if is_debug_enabled():
                            print(f"[DEBUG] Trying fallback: fetch by drive item ID: {uploaded_item['id']}")
                        try:
                            from .graph_api import get_drive_item_with_list_item
                            item_with_list = get_drive_item_with_list_item(
                                site_id, drive_id, uploaded_item['id'],
                                tenant_id, client_id, client_secret, login_endpoint, graph_endpoint
                            )
                            if item_with_list and 'listItem' in item_with_list and 'id' in item_with_list['listItem']:
                                item_id = item_with_list['listItem']['id']
                                if is_debug_enabled():
                                    print(f"[DEBUG] Got list item ID from drive item ID: {item_id}")
                        except Exception as id_fetch_error:
                            if is_debug_enabled():
                                print(f"[DEBUG] Failed to fetch by ID: {str(id_fetch_error)[:200]}")

                if not item_id:
                    # This should rarely happen - we should always be able to query by path
                    print(f"[!] ERROR: Could not get list item ID for {display_path}")
                    print(f"[!] This indicates a critical issue with Graph API access")
                    if is_debug_enabled():
                        print(f"[DEBUG] parent_item_id={parent_item_id}, filename={sanitized_name}")
                        print(f"[DEBUG] uploaded_item keys: {list(uploaded_item.keys()) if uploaded_item else 'None'}")
                    # Don't use filename-only search - it's unreliable for duplicate filenames
                    # The metadata update will be skipped for this file

                if item_id:
                    # Check if we should queue this for batch processing or process immediately
                    if metadata_queue is not None:
                        # Parallel mode: Queue metadata update for batch processing
                        # Store both item_id (for first attempt) and parent_item_id + filename (for retry queries)
                        metadata_queue.put((parent_item_id, sanitized_name, item_id, hash_to_save, is_file_update, display_path))
                        if is_debug_enabled():
                            queue_size = metadata_queue.qsize() if hasattr(metadata_queue, 'qsize') else 'unknown'
                            print(f"[#] Queued FileHash update for {display_path} (queue size: {queue_size})")
                    else:
                        # Sequential mode: Update immediately (backward compatibility)
                        if is_debug_enabled():
                            print(f"[#] Setting FileHash metadata...")

                        debug_metadata = os.environ.get('DEBUG_METADATA', 'false').lower() == 'true'
                        if debug_metadata:
                            print(f"[DEBUG] Setting FileHash for {sanitized_name}")
                            print(f"[DEBUG] SharePoint list item ID: {item_id}")
                            print(f"[DEBUG] About to set FileHash to: {hash_to_save}")

                        # Update the FileHash field using REST API
                        success = update_sharepoint_list_item_field(
                            site_url,
                            list_name,
                            item_id,
                            'FileHash',
                            hash_to_save,
                            tenant_id,
                            client_id,
                            client_secret,
                            login_endpoint,
                            graph_endpoint
                        )

                        if success:
                            if is_debug_enabled():
                                print(f"[✓] FileHash metadata set: {hash_to_save[:8]}...")

                            # Track hash save statistics
                            if is_file_update:
                                upload_stats_dict['hash_updated'] = upload_stats_dict.get('hash_updated', 0) + 1
                            else:
                                upload_stats_dict['hash_new_saved'] = upload_stats_dict.get('hash_new_saved', 0) + 1

                        else:
                            if is_debug_enabled():
                                print(f"[!] Failed to set FileHash metadata via REST API")
                            upload_stats_dict['hash_save_failed'] = upload_stats_dict.get('hash_save_failed', 0) + 1
                else:
                    if is_debug_enabled():
                        print(f"[!] Could not find list item for uploaded file to set hash metadata")
                    # Only track failure in sequential mode (batch mode handles stats after processing)
                    if metadata_queue is None:
                        upload_stats_dict['hash_save_failed'] = upload_stats_dict.get('hash_save_failed', 0) + 1

            except Exception as hash_error:
                print(f"[!] Could not set FileHash metadata via REST API: {str(hash_error)[:200]}")
                # Only track failure in sequential mode (batch mode handles stats after processing)
                if metadata_queue is None:
                    upload_stats_dict['hash_save_failed'] = upload_stats_dict.get('hash_save_failed', 0) + 1
                # Continue anyway - file is uploaded successfully

    except Exception as e:
        upload_stats_dict['failed_files'] += 1
        raise e


def upload_file_with_structure(site_id, drive_id, root_item_id, local_file_path, base_path, site_url, list_name,
                                chunk_size, force_upload, filehash_column_available,
                                tenant_id, client_id, client_secret, login_endpoint,
                                graph_endpoint, upload_stats_dict, max_retry=3, metadata_queue=None,
                                sharepoint_cache=None):
    """
    Upload a file maintaining its directory structure using Graph API.

    Args:
        site_id (str): SharePoint site ID
        drive_id (str): SharePoint drive ID
        root_item_id (str): Root folder item ID where files should be uploaded
        local_file_path (str): The local path of the file to upload
        base_path (str): The base path to strip from the file path (for relative paths)
        site_url (str): Full SharePoint site URL
        list_name (str): Name of the document library (usually "Documents")
        chunk_size (int): Size threshold for using resumable upload
        force_upload (bool): If True, skip comparison and always upload
        filehash_column_available (bool): Whether FileHash column exists in SharePoint
        tenant_id (str): Azure AD tenant ID
        client_id (str): Azure AD app registration client ID
        client_secret (str): Azure AD app registration client secret
        login_endpoint (str): Azure AD authentication endpoint
        graph_endpoint (str): Microsoft Graph API endpoint
        upload_stats_dict (dict): Dictionary to track upload statistics
        max_retry (int): Maximum number of retry attempts (default: 3)
        metadata_queue: Optional BatchQueue for batching metadata updates (parallel mode)
        sharepoint_cache (dict): Optional pre-built cache of SharePoint file metadata
    """
    # Get the relative path of the file
    if base_path:
        rel_path = os.path.relpath(local_file_path, base_path)
    else:
        rel_path = local_file_path

    # Normalize path separators for cross-platform compatibility
    if isinstance(rel_path, bytes):
        rel_path = rel_path.decode('utf-8')
    rel_path = rel_path.replace('\\', '/')

    # Sanitize the entire relative path for SharePoint compatibility
    sanitized_rel_path = sanitize_path_components(rel_path)

    # Get the directory path from sanitized path
    dir_path = os.path.dirname(sanitized_rel_path)

    # Log if path was sanitized
    if sanitized_rel_path != rel_path:
        if is_debug_enabled():
            print(f"[!] Path sanitized for SharePoint: {rel_path} -> {sanitized_rel_path}")

    # Extract folder cache and file cache if available
    # New cache structure: {'files': {...}, 'folders': {...}}
    folder_cache = None
    file_cache = sharepoint_cache  # Default: assume it's the old structure (just files)

    if isinstance(sharepoint_cache, dict) and 'folders' in sharepoint_cache:
        # New structure detected
        folder_cache = sharepoint_cache.get('folders')
        file_cache = sharepoint_cache.get('files')

    # If there's a directory structure, create it in SharePoint
    if dir_path and dir_path != "." and dir_path != "":
        target_folder_id = ensure_folder_exists(
            site_id, drive_id, root_item_id, dir_path,
            tenant_id, client_id, client_secret, login_endpoint, graph_endpoint,
            folder_cache=folder_cache
        )
    else:
        target_folder_id = root_item_id

    # Calculate display path for debug output (use sanitized path as that's what SharePoint sees)
    display_path = sanitized_rel_path

    # Upload the file to the target folder
    if is_debug_enabled():
        print(f"[→] Processing file: {display_path}")
    for i in range(max_retry):
        try:
            upload_file(
                site_id, drive_id, target_folder_id, local_file_path, chunk_size, force_upload,
                site_url, list_name, filehash_column_available,
                tenant_id, client_id, client_secret, login_endpoint,
                graph_endpoint, upload_stats_dict, metadata_queue=metadata_queue,
                display_path=display_path, sharepoint_cache=file_cache
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
