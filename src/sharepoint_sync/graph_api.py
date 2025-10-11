# -*- coding: utf-8 -*-
"""
Microsoft Graph API operations for SharePoint sync.

This module provides all Graph API interactions including column management,
list item operations, and request retry logic.
"""

import time
import requests
from dotenv import load_dotenv
from office365.graph_client import GraphClient
from .auth import acquire_token
from .monitoring import rate_monitor
from .utils import is_debug_mode

# Load environment variables
load_dotenv()

# Global cache for column mappings
column_mapping_cache = {}


def make_graph_request_with_retry(url, headers, method='GET', json_data=None, params=None, max_retries=3):
    """
    Make a Graph API request with proper retry-after handling for 429 responses.
    Includes rate limiting monitoring via response header analysis.

    Args:
        url (str): The Graph API endpoint URL
        headers (dict): Request headers including Authorization
        method (str): HTTP method ('GET', 'POST', 'PATCH', etc.)
        json_data (dict): JSON data for POST/PATCH requests
        params (dict): URL parameters for GET requests
        max_retries (int): Maximum number of retry attempts

    Returns:
        requests.Response: The HTTP response object

    Raises:
        Exception: If all retries are exhausted or non-retryable error occurs
    """
    debug_metadata = is_debug_mode()

    for attempt in range(max_retries + 1):
        try:
            # Add proactive delay if approaching rate limits
            if rate_monitor.should_slow_down() and attempt > 0:
                delay = 2 ** attempt
                print(f"[⚠] Proactive rate limiting delay: {delay}s")
                time.sleep(delay)

            # Make the request based on method
            if method.upper() == 'GET':
                response = requests.get(url, headers=headers, params=params)
            elif method.upper() == 'POST':
                response = requests.post(url, headers=headers, json=json_data)
            elif method.upper() == 'PATCH':
                response = requests.patch(url, headers=headers, json=json_data)
            else:
                raise ValueError(f"Unsupported HTTP method: {method}")

            # Analyze response headers for rate limiting info
            rate_monitor.analyze_response_headers(response)

            # Check for rate limiting (429) or server errors (5xx)
            if response.status_code == 429:
                # Get retry-after header value
                retry_after = response.headers.get('Retry-After', '60')
                try:
                    wait_seconds = int(retry_after)
                except ValueError:
                    wait_seconds = 60  # Default to 60 seconds if header is malformed

                if attempt < max_retries:
                    print(f"[!] Rate limited (429). Waiting {wait_seconds} seconds before retry {attempt + 1}/{max_retries}...")
                    if debug_metadata:
                        print(f"[DEBUG] Retry-After header: {retry_after}")
                        print(f"[DEBUG] Rate limit response: {response.text[:300]}")
                    time.sleep(wait_seconds)
                    continue
                else:
                    print(f"[!] Rate limiting exhausted all retries. Final 429 response:")
                    print(f"[DEBUG] {response.text[:500]}")
                    raise Exception(f"Graph API rate limiting: {response.status_code} after {max_retries} retries")

            elif 500 <= response.status_code < 600:
                # Server error - retry with exponential backoff
                if attempt < max_retries:
                    wait_seconds = (2 ** attempt) + 1  # 1, 3, 7 seconds
                    print(f"[!] Server error ({response.status_code}). Retrying in {wait_seconds} seconds... ({attempt + 1}/{max_retries})")
                    if debug_metadata:
                        print(f"[DEBUG] Server error response: {response.text[:300]}")
                    time.sleep(wait_seconds)
                    continue
                else:
                    print(f"[!] Server errors exhausted all retries. Final response:")
                    print(f"[DEBUG] {response.text[:500]}")
                    raise Exception(f"Graph API server error: {response.status_code} after {max_retries} retries")

            # Success or client error (don't retry client errors like 400, 401, 403, 404)
            return response

        except requests.exceptions.RequestException as e:
            # Network/connection errors - retry with exponential backoff
            if attempt < max_retries:
                wait_seconds = (2 ** attempt) + 1
                print(f"[!] Network error: {e}. Retrying in {wait_seconds} seconds... ({attempt + 1}/{max_retries})")
                time.sleep(wait_seconds)
                continue
            else:
                print(f"[!] Network errors exhausted all retries: {e}")
                raise

    # Should never reach here, but just in case
    raise Exception("Unexpected error in make_graph_request_with_retry")


def get_column_internal_name_mapping(site_id, list_id, token, graph_endpoint):
    """
    Get mapping of display names to internal names for all columns in a SharePoint list.

    Args:
        site_id (str): SharePoint site ID
        list_id (str): SharePoint list/library ID
        token (str): OAuth access token
        graph_endpoint (str): Microsoft Graph API endpoint

    Returns:
        dict: Mapping of display names to column metadata including internal names
              Format: {display_name: {'internal_name': str, 'type': str, 'id': str, 'description': str}}

    Note:
        Results are cached globally in column_mapping_cache to reduce API calls.
    """
    global column_mapping_cache

    # Check cache first
    cache_key = (site_id, list_id)
    if cache_key in column_mapping_cache:
        debug_metadata = is_debug_mode()
        if debug_metadata:
            print(f"[=] Using cached column mappings for site/list")
        return column_mapping_cache[cache_key]

    try:
        debug_metadata = is_debug_mode()

        url = f"https://{graph_endpoint}/v1.0/sites/{site_id}/lists/{list_id}/columns"
        headers = {
            'Authorization': f'Bearer {token}',
            'Accept': 'application/json',
            'Content-Type': 'application/json'
        }

        if debug_metadata:
            print(f"[=] Fetching column mappings from Graph API...")

        response = make_graph_request_with_retry(url, headers, method='GET')

        if response.status_code == 200:
            columns = response.json().get('value', [])
            mapping = {}

            for column in columns:
                display_name = column.get('displayName', '')
                internal_name = column.get('name', '')
                column_type = column.get('columnGroup', 'Unknown')

                mapping[display_name] = {
                    'internal_name': internal_name,
                    'type': column_type,
                    'id': column.get('id', ''),
                    'description': column.get('description', '')
                }

                if debug_metadata:
                    print(f"[=] Column mapping: '{display_name}' -> '{internal_name}' ({column_type})")

            # Cache the result
            column_mapping_cache[cache_key] = mapping

            if debug_metadata:
                print(f"[OK] Cached {len(mapping)} column mappings")

            return mapping
        else:
            print(f"[!] Failed to get column mapping: {response.status_code}")
            if debug_metadata:
                print(f"[DEBUG] Response: {response.text[:500]}")
            return {}

    except Exception as e:
        print(f"[!] Error getting column mapping: {e}")
        return {}


def resolve_field_name(site_id, list_id, token, graph_endpoint, field_name):
    """
    Resolve display name to internal name for reliable field access.

    SharePoint columns have both display names (what users see) and internal names
    (used by API). This function resolves display names to their internal counterparts.

    Args:
        site_id (str): SharePoint site ID
        list_id (str): SharePoint list/library ID
        token (str): OAuth access token
        graph_endpoint (str): Microsoft Graph API endpoint
        field_name (str): Display name or internal name to resolve

    Returns:
        str: The internal name for the field, or original name if not resolved

    Note:
        - Internal names use hex codes for special characters (e.g., '_x0020_' for space)
        - If field_name already appears to be an internal name, returns it as-is
        - Falls back to case-insensitive matching if exact match not found
    """
    try:
        debug_metadata = is_debug_mode()

        # First check if it's already an internal name by checking for hex encoding
        if '_x00' in field_name or (not any(c.isupper() for c in field_name) and '_' in field_name):
            if debug_metadata:
                print(f"[=] '{field_name}' appears to be internal name (contains hex encoding)")
            return field_name

        # Get column mapping
        column_mapping = get_column_internal_name_mapping(site_id, list_id, token, graph_endpoint)

        # Try exact display name match
        if field_name in column_mapping:
            internal_name = column_mapping[field_name]['internal_name']
            if debug_metadata:
                print(f"[OK] Resolved '{field_name}' to internal name '{internal_name}'")
            return internal_name

        # Try case-insensitive match
        for display_name, details in column_mapping.items():
            if display_name.lower() == field_name.lower():
                internal_name = details['internal_name']
                if debug_metadata:
                    print(f"[OK] Resolved '{field_name}' to internal name '{internal_name}' (case-insensitive)")
                return internal_name

        # If no match found, return original name
        if debug_metadata:
            print(f"[!] Could not resolve '{field_name}' to internal name, using as-is")
        return field_name

    except Exception as e:
        print(f"[!] Error resolving field name: {e}")
        return field_name


def sanitize_field_name_for_sharepoint(field_name):
    """
    Convert display name to expected internal name format by encoding special characters.

    SharePoint internal names encode special characters as hex values (e.g., '_x0020_' for space).
    This function attempts to convert a display name to its likely internal name format.

    Args:
        field_name (str): Display name to sanitize

    Returns:
        str: Sanitized field name with special characters encoded

    Note:
        This is a fallback mechanism. Prefer using resolve_field_name() with Graph API
        for accurate internal name resolution.

    Examples:
        'File Hash' -> 'File_x0020_Hash'
        'User#ID' -> 'User_x0023_ID'
        'Value%' -> 'Value_x0025_'
    """
    # Handle common special character conversions
    replacements = {
        ' ': '_x0020_',
        '#': '_x0023_',
        '%': '_x0025_',
        '&': '_x0026_',
        '*': '_x002a_',
        '+': '_x002b_',
        '/': '_x002f_',
        ':': '_x003a_',
        '<': '_x003c_',
        '>': '_x003e_',
        '?': '_x003f_',
        '\\': '_x005c_',
        '|': '_x007c_'
    }

    sanitized = field_name
    for char, replacement in replacements.items():
        sanitized = sanitized.replace(char, replacement)

    return sanitized


def check_and_create_filehash_column(site_url, list_name, tenant_id, client_id, client_secret, login_endpoint, graph_endpoint):
    """
    Check if FileHash column exists in SharePoint document library and create if needed.

    Uses direct Graph API calls to bypass Office365-REST-Python-Client limitations.
    This ensures the FileHash column is available for storing file hashes.

    Args:
        site_url (str): Full SharePoint site URL
        list_name (str): Name of the document library (usually "Documents")
        tenant_id (str): Azure AD tenant ID
        client_id (str): App registration client ID
        client_secret (str): App registration client secret
        login_endpoint (str): Azure AD endpoint
        graph_endpoint (str): Graph API endpoint

    Returns:
        tuple: (success: bool, actual_library_name: str)
               - success: True if column exists or was created successfully
               - actual_library_name: The library name that was actually used (may be fallback)

    Note:
        Requires Sites.ReadWrite.All or Sites.Manage.All permissions.
        The column is created as a single line of text with 255 character limit
        (exact length of xxHash128 hexadecimal representation).
    """
    try:
        # Get token for Graph API
        print("[?] Checking for FileHash column in SharePoint...")
        token = acquire_token(tenant_id, client_id, client_secret, login_endpoint, graph_endpoint)

        if 'access_token' not in token:
            print(f"[!] Failed to acquire token for Graph API: {token.get('error_description', 'Unknown error')}")
            return False, list_name

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
        site_response = make_graph_request_with_retry(site_endpoint, headers, method='GET')

        if site_response.status_code != 200:
            print(f"[!] Failed to get site information: {site_response.status_code}")
            print(f"[DEBUG] Response: {site_response.text[:500]}")
            return False, list_name

        site_data = site_response.json()
        site_id = site_data.get('id')

        if not site_id:
            print("[!] Could not retrieve site ID")
            return False, list_name

        # Get the document library (list) ID
        lists_endpoint = f"https://{graph_endpoint}/v1.0/sites/{site_id}/lists"
        lists_response = make_graph_request_with_retry(lists_endpoint, headers, method='GET')

        if lists_response.status_code != 200:
            print(f"[!] Failed to get lists: {lists_response.status_code}")
            return False, list_name

        lists_data = lists_response.json()
        list_id = None
        actual_library_name = list_name

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
                    actual_library_name = 'Shared Documents'
                    print(f"[!] Using 'Shared Documents' instead of '{list_name}'")
                    break

        if not list_id:
            print(f"[!] Document library '{list_name}' not found")
            return False, list_name

        # Check if FileHash column already exists
        columns_endpoint = f"https://{graph_endpoint}/v1.0/sites/{site_id}/lists/{list_id}/columns"
        columns_response = make_graph_request_with_retry(columns_endpoint, headers, method='GET')

        if columns_response.status_code != 200:
            print(f"[!] Failed to get columns: {columns_response.status_code}")
            return False, actual_library_name

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
                    "maxLength": 255  # xxHash128 produces 32-character hex string
                }
            }

            # Create the column with retry handling
            create_response = make_graph_request_with_retry(
                columns_endpoint,
                headers,
                method='POST',
                json_data=column_definition
            )

            if create_response.status_code == 201:
                print("[✓] FileHash column created successfully")
                # Wait briefly for column to be fully available (eventual consistency)
                time.sleep(2)

                # Verify the newly created column
                is_valid, validation_msg = verify_column_for_filehash_operations(
                    site_id, list_id, token['access_token'], graph_endpoint
                )
                if not is_valid:
                    print(f"[⚠] FileHash column created but verification failed: {validation_msg}")
                    print(f"[⚠] Column may not be immediately accessible (eventual consistency)")
                    # Still return True since column was created, just not immediately accessible

                return True, actual_library_name
            else:
                print(f"[!] Failed to create FileHash column: {create_response.status_code}")
                print(f"[DEBUG] Response: {create_response.text[:500]}")
                return False, actual_library_name

        # Column already exists - verify it's suitable for operations
        is_valid, validation_msg = verify_column_for_filehash_operations(
            site_id, list_id, token['access_token'], graph_endpoint
        )

        if not is_valid:
            print(f"[⚠] FileHash column exists but has issues: {validation_msg}")
            print(f"[⚠] Hash-based comparison may not work correctly")
            # Still return True since column exists, but warn about issues

        return True, actual_library_name

    except Exception as e:
        print(f"[!] Error checking/creating FileHash column: {e}")
        return False, list_name


def rewrite_endpoint(request, graph_endpoint):
    """
    Modify API request URLs for non-standard Microsoft Graph endpoints.

    This function is needed for special Azure environments like:
    - Azure Government Cloud (graph.microsoft.us)
    - Azure Germany (graph.microsoft.de)
    - Azure China (microsoftgraph.chinacloudapi.cn)

    Args:
        request: The HTTP request object to be modified
        graph_endpoint (str): The target Graph API endpoint

    Note:
        This is a callback function used by the GraphClient to intercept
        and modify requests before they're sent.
    """
    # Replace default endpoint with custom one if specified
    request.url = request.url.replace(
        "https://graph.microsoft.com", f"https://{graph_endpoint}"
    )


def get_sharepoint_list_item_by_filename(site_url, list_name, filename, tenant_id, client_id, client_secret, login_endpoint, graph_endpoint):
    """
    Get SharePoint list item by filename using direct Graph API REST calls.

    Args:
        site_url (str): Full SharePoint site URL
        list_name (str): Name of the document library (usually "Documents")
        filename (str): Name of the file to find
        tenant_id (str): Azure AD tenant ID
        client_id (str): App registration client ID
        client_secret (str): App registration client secret
        login_endpoint (str): Azure AD endpoint
        graph_endpoint (str): Graph API endpoint

    Returns:
        dict: List item data with custom columns, or None if not found
    """
    try:
        # Get debug flag
        debug_metadata = is_debug_mode()

        # Get token for Graph API
        token = acquire_token(tenant_id, client_id, client_secret, login_endpoint, graph_endpoint)

        if 'access_token' not in token:
            print(f"[!] Failed to acquire token for Graph API: {token.get('error_description', 'Unknown error')}")
            return None

        headers = {
            'Authorization': f"Bearer {token['access_token']}",
            'Content-Type': 'application/json',
            'Prefer': 'HonorNonIndexedQueriesWarningMayFailRandomly'
        }

        # Parse site URL to get site ID
        site_parts = site_url.replace('https://', '').split('/')
        host_name = site_parts[0]
        site_name = site_parts[2] if len(site_parts) > 2 else ''

        if debug_metadata:
            print(f"[DEBUG] Looking up file: {filename}")
            print(f"[DEBUG] Site parts: host={host_name}, site={site_name}")

        # Get site ID first
        site_endpoint = f"https://{graph_endpoint}/v1.0/sites/{host_name}:/sites/{site_name}"
        site_response = make_graph_request_with_retry(site_endpoint, headers, method='GET')

        if site_response.status_code != 200:
            print(f"[!] Failed to get site information: {site_response.status_code}")
            if debug_metadata:
                print(f"[DEBUG] Site endpoint: {site_endpoint}")
                print(f"[DEBUG] Site response: {site_response.text[:300]}")
            return None

        site_data = site_response.json()
        site_id = site_data.get('id')

        if not site_id:
            print("[!] Could not retrieve site ID")
            if debug_metadata:
                print(f"[DEBUG] Site data keys: {list(site_data.keys())}")
            return None

        if debug_metadata:
            print(f"[DEBUG] Site ID: {site_id}")

        # Get the document library (list) ID
        lists_endpoint = f"https://{graph_endpoint}/v1.0/sites/{site_id}/lists"
        lists_response = make_graph_request_with_retry(lists_endpoint, headers, method='GET')

        if lists_response.status_code != 200:
            print(f"[!] Failed to get lists: {lists_response.status_code}")
            if debug_metadata:
                print(f"[DEBUG] Lists response: {lists_response.text[:300]}")
            return None

        lists_data = lists_response.json()
        list_id = None

        if debug_metadata:
            available_lists = [lst.get('displayName', 'N/A') for lst in lists_data.get('value', [])]
            print(f"[DEBUG] Available lists: {available_lists}")

        for sp_list in lists_data.get('value', []):
            if sp_list.get('displayName') == list_name or sp_list.get('name') == list_name:
                list_id = sp_list.get('id')
                if debug_metadata:
                    print(f"[DEBUG] Found list '{list_name}' with ID: {list_id}")
                break

        if not list_id:
            print(f"[!] Could not find list '{list_name}'")
            return None

        # Check if FileHash column exists in the list before trying to retrieve items
        if debug_metadata:
            print(f"[DEBUG] Checking for FileHash column in list...")
            columns_endpoint = f"https://{graph_endpoint}/v1.0/sites/{site_id}/lists/{list_id}/columns"
            columns_response = make_graph_request_with_retry(columns_endpoint, headers=headers)

            if columns_response.status_code == 200:
                columns_data = columns_response.json()
                filehash_column_found = False
                column_names = []

                for column in columns_data.get('value', []):
                    col_name = column.get('name', 'N/A')
                    col_display_name = column.get('displayName', 'N/A')
                    column_names.append(f"{col_name} ({col_display_name})")

                    if col_name == 'FileHash' or col_display_name == 'FileHash':
                        filehash_column_found = True
                        print(f"[DEBUG] ✓ FileHash column found: name='{col_name}', displayName='{col_display_name}'")
                        print(f"[DEBUG] Column details: {column}")

                if not filehash_column_found:
                    print(f"[DEBUG]  FileHash column NOT found in list")

                print(f"[DEBUG] Available columns: {column_names[:10]}...")  # Show first 10 columns
            else:
                print(f"[DEBUG] Failed to get columns: {columns_response.status_code}")

        # Query list items by filename with expanded fields
        # Try filtering by FileLeafRef first, fallback to getting all items if filter fails
        items_endpoint = f"https://{graph_endpoint}/v1.0/sites/{site_id}/lists/{list_id}/items"

        # First attempt: Use OData filter
        items_params = {
            '$expand': 'fields',
            '$filter': f"fields/FileLeafRef eq '{filename}'"
        }

        if debug_metadata:
            print(f"[DEBUG] Attempting filtered query for: {filename}")
            print(f"[DEBUG] Items endpoint: {items_endpoint}")

        items_response = make_graph_request_with_retry(items_endpoint, headers=headers, params=items_params)

        if items_response.status_code != 200:
            print(f"[!] Failed to get list items with filter: {items_response.status_code}")
            if debug_metadata:
                print(f"[DEBUG] Filter request URL: {items_response.url}")
                print(f"[DEBUG] Response: {items_response.text[:500]}")

            # Fallback: Get all items and filter in Python
            print(f"[DEBUG] Trying fallback: getting all items and filtering in Python...")
            items_params = {'$expand': 'fields'}
            items_response = make_graph_request_with_retry(items_endpoint, headers=headers, params=items_params)

            if items_response.status_code != 200:
                print(f"[!] Failed to get list items (fallback): {items_response.status_code}")
                if debug_metadata:
                    print(f"[DEBUG] Fallback response: {items_response.text[:500]}")
                return None

        items_data = items_response.json()
        items = items_data.get('value', [])

        # Filter items in Python to find matching filename
        if debug_metadata:
            print(f"[DEBUG] Searching through {len(items)} items for '{filename}'")

        for item in items:
            if 'fields' in item and item['fields']:
                file_leaf_ref = item['fields'].get('FileLeafRef')
                if file_leaf_ref == filename:
                    if debug_metadata:
                        print(f"[DEBUG] ✓ Found matching item: {file_leaf_ref}")
                        print(f"[DEBUG] Item ID: {item.get('id', 'N/A')}")
                        print(f"[DEBUG] All available fields in item: {list(item['fields'].keys())}")

                        # Check specifically for FileHash field
                        filehash_value = item['fields'].get('FileHash')
                        if filehash_value:
                            print(f"[DEBUG] ✓ FileHash found in item: {filehash_value}")
                        else:
                            print(f"[DEBUG]  FileHash NOT found in item fields")

                        # Show sample of field values for debugging
                        field_sample = {}
                        for key, value in list(item['fields'].items())[:5]:  # First 5 fields
                            field_sample[key] = str(value)[:50] if value else 'None'
                        print(f"[DEBUG] Sample field values: {field_sample}")

                    return item

        if debug_metadata:
            print(f"[DEBUG]  No matching item found for '{filename}'")
            if items and len(items) > 0:
                sample_names = [item.get('fields', {}).get('FileLeafRef', 'N/A') for item in items[:3]]
                print(f"[DEBUG] Sample FileLeafRef values from list: {sample_names}")

                # Show what fields are available in the first item
                if items[0].get('fields'):
                    sample_fields = list(items[0]['fields'].keys())[:10]
                    print(f"[DEBUG] Sample fields available in first item: {sample_fields}")

        return None

    except Exception as e:
        print(f"[!] Error getting list item by filename: {str(e)[:400]}")
        if is_debug_mode():
            import traceback
            print(f"[DEBUG] Full traceback: {traceback.format_exc()}")
        return None


def update_sharepoint_list_item_field(site_url, list_name, item_id, field_name, field_value, tenant_id, client_id, client_secret, login_endpoint, graph_endpoint):
    """
    Update a custom field in a SharePoint list item using direct Graph API REST calls.

    Args:
        site_url (str): Full SharePoint site URL
        list_name (str): Name of the document library (usually "Documents")
        item_id (str): SharePoint list item ID
        field_name (str): Internal name of the field to update
        field_value (str): Value to set for the field
        tenant_id (str): Azure AD tenant ID
        client_id (str): App registration client ID
        client_secret (str): App registration client secret
        login_endpoint (str): Azure AD endpoint
        graph_endpoint (str): Graph API endpoint

    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Get debug flag
        debug_metadata = is_debug_mode()

        # Get token for Graph API
        token = acquire_token(tenant_id, client_id, client_secret, login_endpoint, graph_endpoint)

        if 'access_token' not in token:
            print(f"[!] Failed to acquire token for Graph API: {token.get('error_description', 'Unknown error')}")
            return False

        headers = {
            'Authorization': f"Bearer {token['access_token']}",
            'Content-Type': 'application/json'
        }

        # Check for rate limiting headers
        if debug_metadata:
            print(f"[DEBUG] Updating field {field_name} = {field_value} for item {item_id}")

        # Parse site URL to get site ID
        site_parts = site_url.replace('https://', '').split('/')
        host_name = site_parts[0]
        site_name = site_parts[2] if len(site_parts) > 2 else ''

        # Get site ID first
        site_endpoint = f"https://{graph_endpoint}/v1.0/sites/{host_name}:/sites/{site_name}"
        site_response = make_graph_request_with_retry(site_endpoint, headers=headers)

        if site_response.status_code != 200:
            print(f"[!] Failed to get site information: {site_response.status_code}")
            if debug_metadata:
                print(f"[DEBUG] Site response: {site_response.text[:300]}")
            return False

        site_data = site_response.json()
        site_id = site_data.get('id')

        if not site_id:
            print("[!] Could not retrieve site ID")
            return False

        # Get the document library (list) ID
        lists_endpoint = f"https://{graph_endpoint}/v1.0/sites/{site_id}/lists"
        lists_response = make_graph_request_with_retry(lists_endpoint, headers=headers)

        if lists_response.status_code != 200:
            print(f"[!] Failed to get lists: {lists_response.status_code}")
            if debug_metadata:
                print(f"[DEBUG] Lists response: {lists_response.text[:300]}")
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

        # Resolve field name to internal name for reliable API access
        resolved_field_name = resolve_field_name(site_id, list_id, token['access_token'], graph_endpoint, field_name)

        if resolved_field_name != field_name and debug_metadata:
            print(f"[=] Resolved field name '{field_name}' to '{resolved_field_name}'")

        # Update the field using PATCH request
        fields_endpoint = f"https://{graph_endpoint}/v1.0/sites/{site_id}/lists/{list_id}/items/{item_id}/fields"
        field_data = {resolved_field_name: field_value}

        if debug_metadata:
            print(f"[DEBUG] PATCH endpoint: {fields_endpoint}")
            print(f"[DEBUG] Field data to update: {field_data}")

        update_response = requests.patch(fields_endpoint, headers=headers, json=field_data)

        # Check for rate limiting headers in response
        if debug_metadata:
            rate_limit_headers = {}
            for header_name, header_value in update_response.headers.items():
                if 'rate' in header_name.lower() or 'throttl' in header_name.lower() or 'limit' in header_name.lower():
                    rate_limit_headers[header_name] = header_value
            if rate_limit_headers:
                print(f"[DEBUG] Rate limiting headers: {rate_limit_headers}")

        if update_response.status_code == 200:
            if debug_metadata:
                print(f"[DEBUG] ✓ Field update successful")
                # Show updated field data
                response_data = update_response.json()
                if field_name in response_data:
                    print(f"[DEBUG] Confirmed field value: {response_data[field_name]}")
            return True
        elif update_response.status_code == 429:
            # Handle throttling specifically
            retry_after = update_response.headers.get('Retry-After', '60')
            print(f"[!] Request throttled (429). Should wait {retry_after} seconds before retry")
            print(f"[DEBUG] Throttling response: {update_response.text[:500]}")
            return False
        else:
            print(f"[!] Failed to update field: {update_response.status_code}")
            print(f"[DEBUG] Response: {update_response.text[:500]}")

            if debug_metadata:
                print(f"[DEBUG] Request headers: {dict(headers)}")
                print(f"[DEBUG] Response headers: {dict(update_response.headers)}")

                # Check if the field name exists
                if update_response.status_code == 400:
                    print(f"[DEBUG] Bad request - field '{field_name}' may not exist or have wrong internal name")

            return False

    except Exception as e:
        print(f"[!] Error updating list item field: {str(e)[:400]}")
        if is_debug_mode():
            import traceback
            print(f"[DEBUG] Full traceback: {traceback.format_exc()}")
        return False


def test_column_accessibility(site_id, list_id, token, graph_endpoint, internal_name):
    """
    Test if a column is accessible by trying to read from list items.

    This function performs a selective query to verify that a column can be accessed
    and read from the SharePoint list. This is useful for detecting columns that exist
    but are not available due to permissions or other restrictions.

    Args:
        site_id (str): SharePoint site ID
        list_id (str): SharePoint list/library ID
        token (str): OAuth access token
        graph_endpoint (str): Microsoft Graph API endpoint
        internal_name (str): Internal name of the column to test

    Returns:
        bool: True if column is accessible, False otherwise

    Note:
        Uses $select to request specific field, which will fail if field is not accessible
    """
    try:
        debug_metadata = is_debug_mode()

        # Try to get list items with specific field selection
        url = f"https://{graph_endpoint}/v1.0/sites/{site_id}/lists/{list_id}/items"
        params = {
            '$top': 1,
            '$expand': 'fields',
            '$select': f'id,fields'
        }
        headers = {
            'Authorization': f'Bearer {token}',
            'Accept': 'application/json'
        }

        if debug_metadata:
            print(f"[=] Testing accessibility of column '{internal_name}'...")

        response = make_graph_request_with_retry(url, headers, params=params)

        if response.status_code == 200:
            # Check if we can access the field in the response
            data = response.json()
            items = data.get('value', [])

            if items and 'fields' in items[0]:
                # Column is accessible if it appears in fields (even if None)
                fields = items[0]['fields']
                if internal_name in fields or any(k.lower() == internal_name.lower() for k in fields.keys()):
                    if debug_metadata:
                        print(f"[OK] Column '{internal_name}' is accessible")
                    return True

            # If no items, assume accessible if query succeeded
            if not items:
                if debug_metadata:
                    print(f"[OK] Column '{internal_name}' query succeeded (no items to verify)")
                return True

            if debug_metadata:
                print(f"[!] Column '{internal_name}' not found in response fields")
            return False
        else:
            if debug_metadata:
                print(f"[!] Column '{internal_name}' accessibility test failed: {response.status_code}")
            return False

    except Exception as e:
        if is_debug_mode():
            print(f"[!] Error testing column accessibility: {e}")
        return False


def comprehensive_column_verification(site_id, list_id, token, graph_endpoint, column_name):
    """
    Comprehensive verification of column existence and properties.

    Performs detailed analysis of a SharePoint column including:
    - Existence verification
    - Property inspection (type, required, indexed, etc.)
    - Accessibility testing
    - Type-specific property analysis

    Args:
        site_id (str): SharePoint site ID
        list_id (str): SharePoint list/library ID
        token (str): OAuth access token
        graph_endpoint (str): Microsoft Graph API endpoint
        column_name (str): Name of column to verify (display or internal name)

    Returns:
        dict: Column analysis dictionary with properties, or None if not found
              Format: {
                  'exists': bool,
                  'display_name': str,
                  'internal_name': str,
                  'id': str,
                  'description': str,
                  'type': str,
                  'required': bool,
                  'hidden': bool,
                  'indexed': bool,
                  'read_only': bool,
                  'enforce_unique': bool,
                  'accessible': bool,
                  'text_properties': dict (if type is text)
              }

    Note:
        Results include detailed property inspection and accessibility testing.
        Use verify_column_for_filehash_operations() for FileHash-specific validation.
    """
    try:
        debug_metadata = is_debug_mode()

        if debug_metadata:
            print(f"[=] Starting comprehensive verification for column '{column_name}'")

        # Step 1: Get all columns with detailed properties
        url = f"https://{graph_endpoint}/v1.0/sites/{site_id}/lists/{list_id}/columns"
        headers = {
            'Authorization': f'Bearer {token}',
            'Accept': 'application/json'
        }

        response = make_graph_request_with_retry(url, headers)

        if response.status_code != 200:
            print(f"[!] Failed to retrieve columns: {response.status_code}")
            if debug_metadata:
                print(f"[DEBUG] Response: {response.text[:500]}")
            return None

        columns_data = response.json().get('value', [])
        target_column = None

        # Step 2: Find target column by name or display name
        for column in columns_data:
            if (column.get('name', '').lower() == column_name.lower() or
                column.get('displayName', '').lower() == column_name.lower()):
                target_column = column
                break

        if not target_column:
            print(f"[!] Column '{column_name}' not found in list")
            if debug_metadata:
                available_columns = [col.get('displayName', col.get('name', 'N/A')) for col in columns_data[:10]]
                print(f"[DEBUG] Available columns (first 10): {available_columns}")
            return None

        # Step 3: Analyze column properties
        column_analysis = {
            'exists': True,
            'display_name': target_column.get('displayName', ''),
            'internal_name': target_column.get('name', ''),
            'id': target_column.get('id', ''),
            'description': target_column.get('description', ''),
            'type': target_column.get('type', ''),
            'required': target_column.get('required', False),
            'hidden': target_column.get('hidden', False),
            'indexed': target_column.get('indexed', False),
            'read_only': target_column.get('readOnly', False),
            'enforce_unique': target_column.get('enforceUniqueValues', False)
        }

        # Step 4: Type-specific analysis
        if 'text' in target_column:
            text_props = target_column['text']
            column_analysis['text_properties'] = {
                'max_length': text_props.get('maxLength', 0),
                'allow_multiple_lines': text_props.get('allowMultipleLines', False),
                'append_changes': text_props.get('appendChangesToExistingText', False)
            }

        # Step 5: Validate column accessibility
        if debug_metadata:
            print(f"[=] Testing column accessibility...")

        accessibility_test = test_column_accessibility(
            site_id, list_id, token, graph_endpoint, column_analysis['internal_name']
        )
        column_analysis['accessible'] = accessibility_test

        # Step 6: Report findings
        if debug_metadata:
            print(f"\n" + "="*40)
            print(f"COLUMN VERIFICATION REPORT")
            print("="*40)
            print(f"Display Name: {column_analysis['display_name']}")
            print(f"Internal Name: {column_analysis['internal_name']}")
            print(f"Type: {column_analysis['type']}")
            print(f"Required: {column_analysis['required']}")
            print(f"Hidden: {column_analysis['hidden']}")
            print(f"Indexed: {column_analysis['indexed']}")
            print(f"Read Only: {column_analysis['read_only']}")
            print(f"Enforce Unique: {column_analysis['enforce_unique']}")

            if 'text_properties' in column_analysis:
                text_props = column_analysis['text_properties']
                print(f"Max Length: {text_props['max_length']}")
                print(f"Multiple Lines: {text_props['allow_multiple_lines']}")

            print(f"Accessible: {column_analysis['accessible']}")

            if column_analysis['hidden']:
                print(f"[⚠] WARNING: Column is hidden")
            if column_analysis['read_only']:
                print(f"[⚠] WARNING: Column is read-only")
            if not column_analysis['accessible']:
                print(f"[!] ERROR: Column exists but is not accessible")

            print("="*40 + "\n")

        return column_analysis

    except Exception as e:
        print(f"[!] Error in comprehensive column verification: {e}")
        if is_debug_mode():
            import traceback
            print(f"[DEBUG] Full traceback: {traceback.format_exc()}")
        return None


def verify_column_for_filehash_operations(site_id, list_id, token, graph_endpoint):
    """
    Specific verification for FileHash column operations.

    Validates that the FileHash column exists, is accessible, and is suitable
    for storing xxHash128 checksums (32-character hex strings).

    Args:
        site_id (str): SharePoint site ID
        list_id (str): SharePoint list/library ID
        token (str): OAuth access token
        graph_endpoint (str): Microsoft Graph API endpoint

    Returns:
        tuple: (is_valid: bool, message: str)
               - is_valid: True if FileHash column is suitable for operations
               - message: Description of validation result or issues found

    Note:
        This function checks:
        - Column existence
        - Column accessibility
        - Column type (must be 'text')
        - Max length (must accommodate 32 characters)
        - Read-only status (must be writable)
        - Hidden status (should not be hidden)
    """
    try:
        debug_metadata = is_debug_mode()

        if debug_metadata:
            print(f"[=] Verifying FileHash column for operations...")

        verification_result = comprehensive_column_verification(
            site_id, list_id, token, graph_endpoint, "FileHash"
        )

        if not verification_result:
            return False, "Column not found"

        # Check if suitable for hash storage
        issues = []

        if verification_result.get('read_only', False):
            issues.append("Column is read-only")

        if verification_result.get('hidden', False):
            issues.append("Column is hidden")

        if not verification_result.get('accessible', False):
            issues.append("Column is not accessible")

        if verification_result.get('type', '') != 'text':
            issues.append(f"Column type is {verification_result.get('type', 'unknown')}, expected 'text'")

        text_props = verification_result.get('text_properties', {})
        max_length = text_props.get('max_length', 0)
        if 0 < max_length < 32:
            issues.append(f"Max length ({max_length}) too small for hash (needs 32)")

        if issues:
            if debug_metadata:
                print(f"[!] FileHash column issues found:")
                for issue in issues:
                    print(f"    - {issue}")
            return False, "; ".join(issues)

        if debug_metadata:
            print(f"[OK] FileHash column is suitable for operations")

        return True, "Column verified successfully"

    except Exception as e:
        error_msg = f"Error during verification: {str(e)[:200]}"
        if is_debug_mode():
            print(f"[!] {error_msg}")
            import traceback
            print(f"[DEBUG] Full traceback: {traceback.format_exc()}")
        return False, error_msg


def create_graph_client(tenant_id, client_id, client_secret, login_endpoint, graph_endpoint):
    """
    Create and configure GraphClient instance for SharePoint access.

    Args:
        tenant_id (str): Azure AD tenant ID
        client_id (str): App registration client ID
        client_secret (str): App registration client secret
        login_endpoint (str): Azure AD endpoint
        graph_endpoint (str): Graph API endpoint

    Returns:
        GraphClient: Configured GraphClient instance
    """
    token_result = acquire_token(tenant_id, client_id, client_secret, login_endpoint, graph_endpoint)

    if 'access_token' not in token_result:
        raise Exception(f"Failed to acquire access token: {token_result.get('error_description', 'Unknown error')}")

    client = GraphClient(lambda: token_result)

    # Apply endpoint rewriting if needed
    if graph_endpoint != "graph.microsoft.com":
        # Create a wrapper that passes graph_endpoint to rewrite_endpoint
        def rewrite_wrapper(request):
            return rewrite_endpoint(request, graph_endpoint)
        client.before_execute(rewrite_wrapper, False)

    return client
