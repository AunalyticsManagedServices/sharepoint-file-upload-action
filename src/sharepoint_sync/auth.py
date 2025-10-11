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

    return token
