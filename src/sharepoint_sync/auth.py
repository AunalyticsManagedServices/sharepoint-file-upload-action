# -*- coding: utf-8 -*-
"""
Microsoft authentication module for SharePoint sync.

This module handles Azure AD authentication using MSAL (Microsoft Authentication Library).
"""

import msal


def acquire_token(tenant_id, client_id, client_secret, login_endpoint, graph_endpoint):
    """
    Acquire an authentication token from Azure Active Directory using MSAL.

    This function handles the OAuth 2.0 client credentials flow, which is used
    for service-to-service authentication (no user interaction required).

    Args:
        tenant_id (str): Azure AD tenant ID (GUID format)
        client_id (str): Application (client) ID from Azure AD app registration
        client_secret (str): Client secret value from Azure AD app registration
        login_endpoint (str): Azure AD authentication endpoint (e.g., 'login.microsoftonline.com')
        graph_endpoint (str): Microsoft Graph API endpoint (e.g., 'graph.microsoft.com')

    Returns:
        dict: Token dictionary containing:
            - 'access_token': The JWT token to authenticate API calls
            - 'token_type': Usually 'Bearer'
            - 'expires_in': Token lifetime in seconds

    Raises:
        Exception: If authentication fails (wrong credentials, network issues, etc.)

    Example:
        token = acquire_token(tenant_id, client_id, client_secret, login_endpoint, graph_endpoint)
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

    # Check for authentication errors (MSAL returns errors in the token dict, not as exceptions)
    if "access_token" not in token:
        error_msg = token.get("error", "unknown_error")
        error_desc = token.get("error_description", "No description provided")
        error_codes = token.get("error_codes", [])

        # Provide user-friendly error messages based on error type
        print("[!] ========================================")
        print("[!] AUTHENTICATION FAILED")
        print("[!] ========================================")

        if "invalid_client" in error_msg or 7000215 in error_codes:
            print("[!] Error: Invalid client credentials")
            print("[!] ")
            print("[!] Troubleshooting steps:")
            print("[!]   1. Verify your CLIENT_ID is correct (check Azure AD app registration)")
            print("[!]   2. Verify your CLIENT_SECRET is correct and hasn't been copied with extra spaces")
            print("[!]   3. Check if the client secret has expired in Azure AD portal")
            print("[!]   4. Ensure you're using the correct TENANT_ID")
            print(f"[!] ")
            print(f"[!] Technical details: {error_desc}")
            raise Exception(f"Authentication failed: Invalid client credentials - {error_desc}")

        elif "unauthorized_client" in error_msg or 700016 in error_codes:
            print("[!] Error: Application not authorized")
            print("[!] ")
            print("[!] Troubleshooting steps:")
            print("[!]   1. Go to Azure AD portal → App registrations → Your app")
            print("[!]   2. Navigate to 'API permissions'")
            print("[!]   3. Verify 'Microsoft Graph' permissions are added:")
            print("[!]      - Sites.ReadWrite.All (minimum)")
            print("[!]      - Sites.Manage.All or Sites.FullControl.All (for FileHash column)")
            print("[!]   4. Click 'Grant admin consent' button (requires admin privileges)")
            print(f"[!] ")
            print(f"[!] Technical details: {error_desc}")
            raise Exception(f"Authentication failed: Application not authorized - {error_desc}")

        elif "invalid_scope" in error_msg or "AADSTS70011" in error_desc:
            print("[!] Error: Invalid scope requested")
            print("[!] ")
            print("[!] Troubleshooting steps:")
            print(f"[!]   1. Verify Graph API endpoint is correct: {graph_endpoint}")
            print("[!]   2. For commercial cloud, use: graph.microsoft.com")
            print("[!]   3. For GovCloud, use: graph.microsoft.us")
            print("[!]   4. For GovCloud High, use: graph.microsoft.us")
            print(f"[!] ")
            print(f"[!] Technical details: {error_desc}")
            raise Exception(f"Authentication failed: Invalid scope - {error_desc}")

        elif "invalid_request" in error_msg:
            print("[!] Error: Invalid authentication request")
            print("[!] ")
            print("[!] Troubleshooting steps:")
            print(f"[!]   1. Verify TENANT_ID format (should be a GUID like: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx)")
            print(f"[!]   2. Verify login endpoint is correct: {login_endpoint}")
            print("[!]   3. For commercial cloud, use: login.microsoftonline.com")
            print("[!]   4. For GovCloud, use: login.microsoftonline.us")
            print(f"[!] ")
            print(f"[!] Technical details: {error_desc}")
            raise Exception(f"Authentication failed: Invalid request - {error_desc}")

        else:
            # Generic error message for unknown error types
            print(f"[!] Error: {error_msg}")
            print("[!] ")
            print("[!] Common issues:")
            print("[!]   - Network connectivity problems")
            print("[!]   - Firewall blocking access to Microsoft identity platform")
            print("[!]   - Incorrect tenant ID or endpoint configuration")
            print(f"[!] ")
            print(f"[!] Technical details:")
            print(f"[!]   Error: {error_msg}")
            print(f"[!]   Description: {error_desc}")
            if error_codes:
                print(f"[!]   Error codes: {error_codes}")
            print("[!] ========================================")
            raise Exception(f"Authentication failed: {error_msg} - {error_desc}")

    return token
