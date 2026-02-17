"""
Microsoft Graph API authentication using MSAL (Public Client flow).
Uses interactive browser auth on first run, then caches refresh tokens.
"""

import json
import os
import sys
import msal

TOKEN_CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "token_cache.json")

SCOPES = [
    "Mail.Read",
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.Read",
    "User.Read",
]


def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_FILE):
        with open(TOKEN_CACHE_FILE, "r") as f:
            cache.deserialize(f.read())
    return cache


def _save_cache(cache: msal.SerializableTokenCache) -> None:
    if cache.has_state_changed:
        with open(TOKEN_CACHE_FILE, "w") as f:
            f.write(cache.serialize())


def get_access_token(client_id: str, tenant_id: str) -> str:
    """
    Acquire a valid Microsoft Graph access token.
    On first run, opens a browser for interactive login.
    Subsequent runs use the cached refresh token silently.
    """
    cache = _load_cache()
    authority = f"https://login.microsoftonline.com/{tenant_id}"

    app = msal.PublicClientApplication(
        client_id,
        authority=authority,
        token_cache=cache,
    )

    # Try silent acquisition first (uses cached refresh token)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_cache(cache)
            return result["access_token"]

    # Fall back to interactive login (opens browser)
    print("No cached token found. Opening browser for Microsoft login...", file=sys.stderr)
    result = app.acquire_token_interactive(
        scopes=SCOPES,
        prompt="select_account",
    )

    if "access_token" not in result:
        error = result.get("error_description", result.get("error", "Unknown error"))
        raise RuntimeError(f"Failed to acquire token: {error}")

    _save_cache(cache)
    return result["access_token"]
