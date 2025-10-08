import sys
import os
import msal
import glob
import time
from office365.graph_client import GraphClient
from office365.runtime.odata.v4.upload_session_request import UploadSessionRequest
from office365.onedrive.driveitems.driveItem import DriveItem
from office365.onedrive.internal.paths.url import UrlPath
from office365.runtime.queries.upload_session import UploadSessionQuery
from office365.onedrive.driveitems.uploadable_properties import DriveItemUploadableProperties

site_name = sys.argv[1]
sharepoint_host_name = sys.argv[2]
tenant_id = sys.argv[3]
client_id = sys.argv[4]
client_secret = sys.argv[5]
upload_path = sys.argv[6]
file_path = sys.argv[7]
max_retry = int(sys.argv[8]) or 3
login_endpoint = sys.argv[9] or "login.microsoftonline.com"
graph_endpoint = sys.argv[10] or "graph.microsoft.com"
file_path_recursive_match = sys.argv[11] if len(sys.argv) > 11 and sys.argv[11] else "False"

# below used with 'get_by_url' in GraphClient calls
tenant_url = f'https://{sharepoint_host_name}/sites/{site_name}'

# Convert string to boolean for recursive flag
recursive = file_path_recursive_match.lower() in ['true', '1', 'yes']

# Get all files and directories matching the pattern
local_items = glob.glob(file_path, recursive=recursive)

if not local_items:
    print(f"[Error] No files or directories matched pattern: {file_path}")
    sys.exit(1)

# Separate files and directories
local_files = []
local_dirs = []

for item in local_items:
    if os.path.isfile(item):
        local_files.append(item)
    elif os.path.isdir(item):
        local_dirs.append(item)
        # If a directory is found, also get all files within it recursively
        for root, dirs, files in os.walk(item):
            for file in files:
                local_files.append(os.path.join(root, file))

if not local_files and not local_dirs:
    print(f"[Error] No files or directories found matching pattern: {file_path}")
    sys.exit(1)

print(f"Found {len(local_files)} file(s) and {len(local_dirs)} directory(ies) to process")

def acquire_token():
    """
    Acquire token via MSAL
    """
    authority_url = f'https://{login_endpoint}/{tenant_id}'
    app = msal.ConfidentialClientApplication(
        authority=authority_url,
        client_id=client_id,
        client_credential=client_secret
    )
    token = app.acquire_token_for_client(scopes=[f"https://{graph_endpoint}/.default"])
    return token

#Replace office365 request url with the correct endpoint for non-default environments
def rewrite_endpoint(request):
    request.url = request.url.replace(
        "https://graph.microsoft.com", f"https://{graph_endpoint}"
    )

client = GraphClient(acquire_token)
client.before_execute(rewrite_endpoint, False)
root_drive = client.sites.get_by_url(tenant_url).drive.root.get_by_path(upload_path)

# Cache for created folders to avoid recreating them
created_folders = {}

def ensure_folder_exists(parent_drive, folder_path):
    """
    Recursively ensure that a folder structure exists in SharePoint
    Returns the DriveItem for the final folder
    """
    # If we've already created this folder, return it from cache
    if folder_path in created_folders:
        return created_folders[folder_path]
    
    # Split the path into components
    path_parts = folder_path.split(os.sep)
    current_drive = parent_drive
    current_path = ""
    
    for folder_name in path_parts:
        if not folder_name:  # Skip empty parts
            continue
            
        # Build the current path
        if current_path:
            current_path = f"{current_path}/{folder_name}"
        else:
            current_path = folder_name
            
        # Check if we've already created this path
        if current_path in created_folders:
            current_drive = created_folders[current_path]
            continue
        
        try:
            # Try to get the folder
            folder = current_drive.get_by_path(folder_name).get().execute_query()
            current_drive = folder
            created_folders[current_path] = folder
            print(f"[✓] Folder exists: {current_path}")
        except Exception as e:
            # Folder doesn't exist, create it
            try:
                print(f"[+] Creating folder: {current_path}")
                folder = current_drive.children.add(folder_name, folder=True).execute_query()
                current_drive = folder
                created_folders[current_path] = folder
                print(f"[✓] Created folder: {current_path}")
            except Exception as create_error:
                print(f"[Error] Failed to create folder {current_path}: {create_error}")
                raise create_error
    
    return current_drive

def progress_status(offset, file_size):
    print(f"Uploaded {offset} bytes from {file_size} bytes ... {offset/file_size*100:.2f}%")

def success_callback(remote_file, local_path):
    print(f"[✓] File {local_path} has been uploaded to {remote_file.web_url}")

def resumable_upload(drive, local_path, file_size, chunk_size, max_chunk_retry, timeout_secs):
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
    
    file_name = os.path.basename(local_path)
    return_type = DriveItem(
        drive.context, 
        UrlPath(file_name, drive.resource_path))
    qry = UploadSessionQuery(
        return_type, {"item": DriveItemUploadableProperties(name=file_name)})
    drive.context.add_query(qry).after_query_execute(_start_upload)
    return_type.get().execute_query()
    success_callback(return_type, local_path)

def upload_file(drive, local_path, chunk_size):
    file_size = os.path.getsize(local_path)
    if file_size < chunk_size:
        remote_file = drive.upload_file(local_path).execute_query()
        success_callback(remote_file, local_path)
    else:
        resumable_upload(
            drive, 
            local_path, 
            file_size, 
            chunk_size, 
            max_chunk_retry=60, 
            timeout_secs=10*60)

def upload_file_with_structure(root_drive, local_file_path, base_path=""):
    """
    Upload a file maintaining its directory structure
    
    :param root_drive: The root drive in SharePoint where files should be uploaded
    :param local_file_path: The local path of the file to upload
    :param base_path: The base path to strip from the file path (for relative paths)
    """
    # Get the relative path of the file
    if base_path:
        rel_path = os.path.relpath(local_file_path, base_path)
    else:
        rel_path = local_file_path
    
    # Get the directory path and file name
    dir_path = os.path.dirname(rel_path)
    file_name = os.path.basename(rel_path)
    
    # If there's a directory structure, create it in SharePoint
    if dir_path and dir_path != ".":
        target_folder = ensure_folder_exists(root_drive, dir_path)
    else:
        target_folder = root_drive
    
    # Upload the file to the target folder
    print(f"[→] Uploading {local_file_path} to SharePoint...")
    for i in range(max_retry):
        try:
            upload_file(target_folder, local_file_path, 4*1024*1024)
            break
        except Exception as e:
            print(f"Unexpected error occurred: {e}, {type(e)}")
            if i == max_retry - 1:
                raise e
            else:
                print(f"Retrying... ({i+1}/{max_retry})")
                time.sleep(2)

# Determine the base path for relative path calculation
# This helps maintain the correct directory structure in SharePoint
base_path = ""
if local_dirs:
    # If we have directories, use the parent of the first directory as base
    base_path = os.path.dirname(local_dirs[0])
elif local_files:
    # If we only have files, use their common parent directory
    base_path = os.path.dirname(os.path.commonpath(local_files))

# Upload all files with their directory structure
for f in local_files:
    if os.path.isfile(f):  # Double-check it's a file
        upload_file_with_structure(root_drive, f, base_path)
    else:
        print(f"[Warning] Skipping {f} as it's not a file")

print(f"[✓] Upload process completed. Processed {len(local_files)} file(s)")
