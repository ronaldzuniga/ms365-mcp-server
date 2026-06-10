"""
Microsoft Graph API authentication using MSAL (Public Client flow).
Uses interactive browser auth on first run, then caches refresh tokens.
"""

import os
import sys
from typing import Optional

import msal

TOKEN_CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "token_cache.json")

SCOPES = [
    "Mail.Read",
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.Read",
    "User.Read",
]

# Keep one MSAL app (and its token cache) alive for the life of the process so
# every token request can reuse the cached access token or refresh it silently.
_app: Optional[msal.PublicClientApplication] = None
_cache: Optional[msal.SerializableTokenCache] = None


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


def _get_app(client_id: str, tenant_id: str) -> msal.PublicClientApplication:
    global _app, _cache
    if _app is None:
        _cache = _load_cache()
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        _app = msal.PublicClientApplication(
            client_id,
            authority=authority,
            token_cache=_cache,
        )
    return _app


def get_access_token(
    client_id: str,
    tenant_id: str,
    force_refresh: bool = False,
    allow_interactive: bool = True,
) -> str:
    """
    Acquire a valid Microsoft Graph access token.

    Tries the in-memory/disk cache first: MSAL returns the cached access token
    while it is still valid and silently exchanges the refresh token when it
    has expired, so this is safe (and cheap) to call before every request.

    Args:
        force_refresh: Skip the cached access token and force a refresh-token
            exchange (use after a 401 from Graph).
        allow_interactive: If silent acquisition fails, open a browser login.
            Set to False inside tool calls so the server never hangs waiting
            for a browser that the user may not see.
    """
    app = _get_app(client_id, tenant_id)

    # Try silent acquisition first (uses cached access/refresh token)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(
            SCOPES, account=accounts[0], force_refresh=force_refresh
        )
        if result and "access_token" in result:
            _save_cache(_cache)
            return result["access_token"]

    if not allow_interactive:
        raise RuntimeError(
            "Microsoft authentication expired and could not be refreshed silently. "
            "Run 'python server.py' once to sign in again (or delete token_cache.json first)."
        )

    # Fall back to interactive login (opens browser)
    print("No cached token found. Opening browser for Microsoft login...", file=sys.stderr)
    result = app.acquire_token_interactive(
        scopes=SCOPES,
        prompt="select_account",
    )

    if "access_token" not in result:
        error = result.get("error_description", result.get("error", "Unknown error"))
        raise RuntimeError(f"Failed to acquire token: {error}")

    _save_cache(_cache)
    return result["access_token"]
