# Microsoft 365 MCP Server

An MCP server that gives Claude Desktop and Claude Code access to your Microsoft 365 email (Outlook) and calendar via Microsoft Graph API.

## Quick Start Overview

Setting up this MCP server requires two main steps:

1. **Azure AD Setup** (5-10 minutes, one-time): Create an app registration in Azure to get your credentials
2. **Local Installation** (5 minutes): Install the server and configure Claude Desktop or Claude Code

Both are detailed below with step-by-step instructions.

## Tools Included

| Tool | Description |
|------|-------------|
| `ms365_list_emails` | List emails from any folder (inbox, sent, etc.) |
| `ms365_read_email` | Read full email content by ID |
| `ms365_search_emails` | Search emails by keyword |
| `ms365_send_email` | Send an email |
| `ms365_create_draft` | Create a draft email without sending |
| `ms365_reply_email` | Reply to an email |
| `ms365_list_mail_folders` | List all mail folders with counts |
| `ms365_list_events` | List calendar events in a date range |
| `ms365_get_event` | Get full event details |
| `ms365_list_calendars` | List all calendars |
| `ms365_find_free_time` | Check free/busy for people |
| `ms365_get_profile` | Get your Microsoft 365 profile |

---

## Azure AD Setup (Required First)

Before you can use this MCP server, you need to create an Azure AD App Registration to get your `MS_CLIENT_ID` and `MS_TENANT_ID`. This only needs to be done once.

### Step 1: Create the App Registration

1. **Go to the Azure Portal**
   - Navigate to [portal.azure.com](https://portal.azure.com)
   - Sign in with your Microsoft 365 account

2. **Access App Registrations**
   - In the top search bar, type "App registrations"
   - Click on **App registrations** from the search results

3. **Create New Registration**
   - Click **"+ New registration"** button at the top
   - Fill in the registration form:
     - **Name**: `MS365 MCP Server` (or any name you prefer)
     - **Supported account types**: Select **"Accounts in this organizational directory only (Single tenant)"**
     - **Redirect URI**:
       - Select **"Mobile and desktop applications"** from the dropdown
       - Enter: `http://localhost`
   - Click **"Register"** button

4. **Copy Your Credentials**
   - You'll be taken to the app's Overview page
   - Copy and save these two values (you'll need them later):
     - **Application (client) ID** → This is your `MS_CLIENT_ID`
     - **Directory (tenant) ID** → This is your `MS_TENANT_ID`
   - Keep these values handy — you'll use them in the setup steps below

### Step 2: Set API Permissions

1. **Navigate to API Permissions**
   - In the left sidebar, click **"API permissions"**

2. **Add Microsoft Graph Permissions**
   - Click **"+ Add a permission"**
   - Select **"Microsoft Graph"**
   - Select **"Delegated permissions"**

3. **Add Required Permissions**
   - Search for and add each of these permissions:
     - ✅ **Mail.Read** - Read your emails
     - ✅ **Mail.ReadWrite** - Create and manage email drafts
     - ✅ **Mail.Send** - Send emails from your account
     - ✅ **Calendars.Read** - Read your calendar
     - ✅ **User.Read** - Read basic profile information
   - Click **"Add permissions"** when done

4. **Grant Admin Consent**
   - Back on the API permissions page, look for the **"Grant admin consent for [Your Organization]"** button
   - **If you're an admin**: Click the button and confirm
   - **If you're not an admin**: Ask your IT administrator to grant consent
   - You should see green checkmarks under the "Status" column for all permissions

### Step 3: Verify Your Setup

Your Azure configuration is complete when you have:
- ✅ Application (client) ID copied
- ✅ Directory (tenant) ID copied
- ✅ All five permissions added and granted (green checkmarks)

Now you're ready to proceed with the local installation!

---

## Local Installation

### Prerequisites

- Python 3.10+
- Claude Desktop or Claude Code CLI installed
- **MS_CLIENT_ID** and **MS_TENANT_ID** from Azure setup above

### Step 1: Clone or Download the Project

Copy the `ms365-mcp-server` folder to a permanent location on your machine:

```bash
# Example: put it in your home directory
cp -r ms365-mcp-server ~/ms365-mcp-server
cd ~/ms365-mcp-server
```

### Step 2: Create a Virtual Environment and Install Dependencies

```bash
cd ~/ms365-mcp-server
python3 -m venv .venv
source .venv/bin/activate       # macOS/Linux
# .venv\Scripts\activate        # Windows

pip install -r requirements.txt
```

### Step 3: First-Time Authentication

Run the server once manually to complete the interactive Microsoft login:

```bash
cd ~/ms365-mcp-server
source .venv/bin/activate

export MS_CLIENT_ID="your-client-id-here"
export MS_TENANT_ID="your-tenant-id-here"

python server.py
```

This will open your browser for Microsoft login. Sign in and grant permissions.
Once done, a `token_cache.json` file is created — this stores your refresh token so you won't need to log in again (tokens auto-renew for ~90 days).

Press `Ctrl+C` to stop the server after the login succeeds.

### Step 4: Configure Claude Desktop

Open Claude Desktop's configuration file:

**IMPORTANT**: Use the `MS_CLIENT_ID` and `MS_TENANT_ID` you copied from the Azure setup above.

- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
- **Linux**: `~/.config/Claude/claude_desktop_config.json`

Add the MCP server configuration. Replace the paths and credentials with your own:

```json
{
  "mcpServers": {
    "ms365": {
      "command": "/Users/YOURUSER/ms365-mcp-server/.venv/bin/python",
      "args": [
        "/Users/YOURUSER/ms365-mcp-server/server.py"
      ],
      "env": {
        "MS_CLIENT_ID": "your-client-id-here",
        "MS_TENANT_ID": "your-tenant-id-here"
      }
    }
  }
}
```

> **Important:** Use the FULL path to the Python binary inside your virtual environment, not just `python3`.

**Finding your Python path:**
```bash
# macOS/Linux
source ~/ms365-mcp-server/.venv/bin/activate
which python
# Outputs something like: /Users/youruser/ms365-mcp-server/.venv/bin/python

# Windows
# Usually: C:\Users\YOURUSER\ms365-mcp-server\.venv\Scripts\python.exe
```

### Step 5: Restart Claude Desktop

Quit Claude Desktop completely and reopen it. You should see a 🔌 icon or the MCP tools available when you start a conversation.

---

## Setup for Claude Code (CLI)

If you're using Claude Code (the command-line tool), follow these steps:

### Prerequisites
- Python 3.10+
- Claude Code CLI installed
- An Azure AD App Registration (MS_CLIENT_ID and MS_TENANT_ID)

### Step 1-3: Same as Claude Desktop
Follow Steps 1-3 from the Claude Desktop setup above (clone, create venv, first-time auth).

### Step 4: Configure Claude Code MCP Settings

Open or create the Claude Code MCP configuration file:

**IMPORTANT**: Use the `MS_CLIENT_ID` and `MS_TENANT_ID` you copied from the Azure setup above.
- **macOS/Linux**: `~/.claude/mcp_settings.json`
- **Windows**: `%USERPROFILE%\.claude\mcp_settings.json`

Add the MCP server configuration:

```json
{
  "mcpServers": {
    "ms365": {
      "command": "/Users/YOURUSER/ms365_mcp_server/.venv/bin/python",
      "args": [
        "/Users/YOURUSER/ms365_mcp_server/server.py"
      ],
      "env": {
        "MS_CLIENT_ID": "your-client-id-here",
        "MS_TENANT_ID": "your-tenant-id-here"
      }
    }
  }
}
```

**Important:** Use the FULL path to the Python binary inside your virtual environment.

### Step 5: Test the Connection

Start Claude Code and verify the MCP server is loaded:
```bash
claude-code
```

The ms365 tools should be available in your session.

---

## Usage Examples

Once configured, just chat naturally with Claude Desktop:

- "Read my latest emails"
- "Search my emails for messages from Mario about Milwaukee Tool"
- "What's on my calendar this week?"
- "Find free time between me and mario@qualitara.com tomorrow"
- "Send an email to john@example.com with subject 'Meeting Notes' saying..."
- "Reply to that last email saying I'll be there at 3pm"

---

## Troubleshooting

### "Authentication expired" errors
Delete `token_cache.json` and run `python server.py` again to re-authenticate.

### Tools not showing in Claude Desktop
1. Check the config file path is correct
2. Make sure the Python path points to the venv's python (not system python)
3. Restart Claude Desktop completely (quit, not just close window)
4. Check Claude Desktop logs for errors

### Permission errors from Microsoft Graph
Go back to your Azure AD App Registration → API Permissions and make sure `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`, `Calendars.Read`, and `User.Read` are all granted with admin consent.

### Token cache location
The `token_cache.json` is saved in the same directory as `server.py`. Make sure the process has write access to that folder.

---

## File Structure

```
ms365-mcp-server/
├── server.py           # Main MCP server with all tools
├── auth.py             # MSAL authentication (token acquire + cache)
├── graph_client.py     # Microsoft Graph API client
├── requirements.txt    # Python dependencies
├── token_cache.json    # Auto-generated after first login (DO NOT COMMIT)
└── README.md           # This file
```

---

## Security Notes

- **No client secret is needed.** This uses MSAL's `PublicClientApplication` with interactive auth — the Microsoft-recommended approach for desktop/CLI apps.
- `token_cache.json` contains your refresh token. Treat it like a password. Do not commit it to version control.
- The refresh token auto-renews for ~90 days on each use.
