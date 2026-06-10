# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an MCP (Model Context Protocol) server that provides Claude Desktop with access to Microsoft 365 email (Outlook) and calendar via Microsoft Graph API. It uses MSAL (Microsoft Authentication Library) with PublicClientApplication for interactive browser-based authentication.

## Core Architecture

The codebase is organized into three main modules:

1. **server.py** - MCP server using FastMCP framework
   - Defines 10 tools exposed to Claude Desktop (email and calendar operations)
   - Uses Pydantic models for input validation
   - All tools call Graph through `_call_graph()`, which acquires a valid token via `_get_token()` before every request (MSAL serves it from cache while valid, silently refreshes when expired) and retries once with a forced refresh on 401
   - Consistent error handling via `_handle_error()` that interprets HTTP status codes

2. **auth.py** - MSAL authentication layer
   - Uses a single long-lived `msal.PublicClientApplication` (no client secret required)
   - Caches tokens in `token_cache.json` (auto-created, never commit)
   - `get_access_token()` tries silent refresh first, falls back to interactive browser login (only when `allow_interactive=True`; tool calls pass `False` so they never hang waiting for a browser)
   - Scopes: `Mail.Read`, `Mail.Send`, `Calendars.Read`, `User.Read`

3. **graph_client.py** - HTTP client for Microsoft Graph API
   - All functions are async using `httpx.AsyncClient`
   - Base URL: `https://graph.microsoft.com/v1.0`
   - Generic `graph_request()` wrapper handles all HTTP methods
   - Endpoint functions (e.g., `list_messages`, `send_message`) build params and call `graph_request()`

## Environment Variables

Required for operation:
- `MS_CLIENT_ID` - Azure AD App Registration client ID
- `MS_TENANT_ID` - Azure AD tenant ID

These are typically set in Claude Desktop's `claude_desktop_config.json` under the `env` section of the MCP server configuration.

## Development Commands

### Setup
```bash
# Create and activate virtual environment
python3 -m venv .venv
source .venv/bin/activate  # macOS/Linux
# .venv\Scripts\activate   # Windows

# Install dependencies
pip install -r requirements.txt
```

### Running the Server

**Interactive authentication (first time or to refresh expired tokens):**
```bash
source .venv/bin/activate
export MS_CLIENT_ID="your-client-id"
export MS_TENANT_ID="your-tenant-id"
python server.py
```

This opens a browser for Microsoft login and creates `token_cache.json`.

**Normal usage:**
The server is launched automatically by Claude Desktop using the configuration in `claude_desktop_config.json`. It does not run standalone for production use.

### Testing During Development

Since there are no unit tests, manual testing is done by:
1. Running `python server.py` directly
2. Testing tools in Claude Desktop after configuration
3. Monitoring stderr output for authentication/error messages

## Important Constraints

- **Token refresh**: Access tokens auto-refresh on every tool call via MSAL's cached refresh token (valid ~90 days with use); no server restart needed when the access token expires
- **No client secret**: Uses interactive auth flow, not client credentials
- **Token cache location**: Always saved in same directory as `server.py`
- **Timeout**: All Graph API requests have a 30-second timeout
- **Email pagination**: `ms365_list_emails` has max limit of 50 emails per call
- **Search limits**: `ms365_search_emails` max 25 results per query

## Microsoft Graph API Patterns

All Graph API operations follow this pattern:
1. Call the async function in `graph_client.py` via `_call_graph()` (which acquires/refreshes the token and retries once on 401)
2. Parse JSON response and format as markdown
3. Catch exceptions and use `_handle_error()` for consistent error messages

When adding new tools:
- Add Pydantic input model with validation
- Define tool with `@mcp.tool()` decorator and appropriate annotations
- Call graph_client function
- Format response as markdown
- Use `_handle_error()` for exception handling

## Authentication Flow

1. Server starts → `_get_token(allow_interactive=True)` is called once
2. MSAL checks `token_cache.json` for a cached account
3. If cached token exists → silently refreshes; if no cache → opens browser for interactive login
4. Every tool call re-acquires the token via `_call_graph()` → MSAL returns the cached access token while valid and silently exchanges the refresh token when it expires (~1 hour), so long-running sessions keep working without restarting Claude Desktop
5. If Graph still returns 401 → one retry with `force_refresh=True`
6. Only if the refresh token itself is dead (~90 days unused or revoked) does the user need to run `python server.py` to sign in again

## Troubleshooting

**"Authentication expired" errors:**
Delete `token_cache.json` and run `python server.py` to re-authenticate.

**Tools not appearing in Claude Desktop:**
- Verify Python path in config points to venv's python (not system python)
- Check Claude Desktop logs
- Restart Claude Desktop completely (quit, not just close window)

**Permission errors:**
Ensure Azure AD App Registration has `Mail.Read`, `Mail.Send`, `Calendars.Read`, `User.Read` permissions granted with admin consent.
