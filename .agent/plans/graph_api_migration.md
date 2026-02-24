# InboxBar: Microsoft Graph API Migration Plan
## Making InboxBar work with New Outlook / Office 365

---

## 1. Overview

### Current State
InboxBar uses **COM automation** (`win32com.client`) to talk to the **classic Outlook desktop app** directly. This works great but is limited to machines running the traditional Win32 Outlook.

### Target State
Add a **second backend** using the **Microsoft Graph REST API** so InboxBar works with:
- ✅ Classic Outlook (existing COM backend — unchanged)
- ✅ New Outlook for Windows (web-based, no COM support)
- ✅ Users without Outlook installed (just an M365 account)
- ✅ Outlook on the web (OWA)

### Architecture: Dual Backend
```
sidebar_main.py
      │
      ▼
  MailClient (Abstract Interface)
      │
      ├── OutlookCOMClient (existing, for classic Outlook)
      │     └── win32com.client
      │
      └── GraphAPIClient (new, for M365/New Outlook)
            └── Microsoft Graph REST API via MSAL + requests
```

---

## 2. Prerequisites & Azure Setup

### 2.1 Azure AD App Registration
1. Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
2. Click **New registration**
   - Name: `InboxBar Desktop`
   - Supported account types: **Accounts in any organizational directory and personal Microsoft accounts**
   - Redirect URI: **Public client/native** → `http://localhost:8400`
3. Note down:
   - **Application (client) ID** — this goes in the app config
   - **Tenant ID** — use `common` for multi-tenant

### 2.2 API Permissions Required
| Permission | Type | Purpose |
|---|---|---|
| `Mail.ReadWrite` | Delegated | Read/write emails, mark read, delete |
| `Mail.Send` | Delegated | Send emails (quick create, share) |
| `Calendars.ReadWrite` | Delegated | Read calendar events, create meetings |
| `Tasks.ReadWrite` | Delegated | Read/complete To Do tasks |
| `User.Read` | Delegated | Get user profile (account name) |
| `offline_access` | Delegated | Refresh tokens (stay logged in) |

### 2.3 New Python Dependencies
```
msal>=1.25.0        # Microsoft Authentication Library
requests>=2.31.0    # HTTP client for Graph API calls
```
Add to `requirements.txt` and PyInstaller hidden imports.

---

## 3. Implementation Phases

### Phase 1: Abstract Interface & Backend Selection (Day 1, ~2 hours)

#### 3.1 Create Abstract Mail Client Interface
**File: `sidebar/services/mail_client.py`**

```python
from abc import ABC, abstractmethod

class MailClient(ABC):
    """Abstract interface that both COM and Graph backends implement."""
    
    # --- Connection ---
    @abstractmethod
    def connect(self) -> bool: ...
    
    @abstractmethod
    def reconnect(self) -> bool: ...
    
    @abstractmethod
    def is_connected(self) -> bool: ...
    
    # --- Accounts ---
    @abstractmethod
    def get_accounts(self) -> list: ...
    
    # --- Email ---
    @abstractmethod
    def get_inbox_items(self, count=20, unread_only=False, only_flagged=False, 
                        due_filters=None, account_names=None, account_config=None) -> tuple: ...
    
    @abstractmethod
    def get_unread_count(self, account_names=None, account_config=None) -> int: ...
    
    @abstractmethod
    def mark_as_read(self, entry_id, store_id=None) -> bool: ...
    
    @abstractmethod
    def delete_email(self, entry_id, store_id=None) -> bool: ...
    
    @abstractmethod
    def toggle_flag(self, entry_id, store_id=None) -> bool: ...
    
    @abstractmethod
    def unflag_email(self, entry_id, store_id=None) -> bool: ...
    
    @abstractmethod
    def open_item(self, entry_id, store_id=None): ...
    
    # --- Calendar ---
    @abstractmethod
    def get_calendar_items(self, start_dt, end_dt, account_names=None) -> list: ...
    
    # --- Tasks ---
    @abstractmethod
    def get_tasks(self, due_filters=None, account_names=None) -> list: ...
    
    @abstractmethod
    def mark_task_complete(self, entry_id, store_id=None) -> bool: ...
    
    # --- Quick Create ---
    @abstractmethod
    def create_email(self): ...
    
    @abstractmethod
    def create_meeting(self): ...
    
    @abstractmethod
    def create_task(self): ...
    
    # --- New Mail Detection ---
    @abstractmethod
    def check_new_mail(self, account_names=None) -> bool: ...
    
    # --- Pulse ---
    @abstractmethod
    def get_pulse_status(self, account_names=None) -> dict: ...
    
    # --- Utility ---
    @abstractmethod
    def get_category_map(self) -> dict: ...
    
    @abstractmethod
    def search_contacts(self, query, max_results=8) -> list: ...
    
    @abstractmethod
    def get_folder_list(self, account_name=None) -> list: ...
```

#### 3.2 Make Existing OutlookClient implement MailClient
**File: `sidebar/services/outlook_client.py`**

- Add `from sidebar.services.mail_client import MailClient`
- Change class declaration to `class OutlookClient(MailClient):`
- Rename internal helper `_is_connection_healthy` → expose as `is_connected`
- Existing methods already match the interface — minimal changes

#### 3.3 Add Backend Selection to Config
**File: `sidebar/core/config_manager.py`**

```python
self.backend = "auto"  # "auto" | "com" | "graph"
```

**Auto-detection logic** in `sidebar_main.py`:
```python
def _select_backend(self):
    if self.config.backend == "com":
        return OutlookClient()
    elif self.config.backend == "graph":
        return GraphAPIClient()
    else:  # "auto"
        # Try COM first (classic Outlook), fall back to Graph
        try:
            client = OutlookClient()
            if client.is_connected():
                return client
        except:
            pass
        return GraphAPIClient()
```

---

### Phase 2: Authentication & Token Management (Day 1, ~3 hours)

#### 3.4 MSAL Token Manager
**File: `sidebar/services/graph_auth.py`**

This handles the OAuth2 flow and token caching.

```python
import msal
import os
import json

CLIENT_ID = "YOUR_APP_ID_HERE"  # From Azure AD app registration
SCOPES = [
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.ReadWrite",
    "Tasks.ReadWrite",
    "User.Read",
]

class GraphAuth:
    def __init__(self):
        cache_path = os.path.join(os.environ.get("LOCALAPPDATA", "."), "InboxBar", "graph_token_cache.bin")
        self.cache = msal.SerializableTokenCache()
        
        if os.path.exists(cache_path):
            self.cache.deserialize(open(cache_path, "r").read())
        
        self.app = msal.PublicClientApplication(
            CLIENT_ID,
            authority="https://login.microsoftonline.com/common",
            token_cache=self.cache,
        )
        self._cache_path = cache_path
    
    def get_token(self):
        """Returns a valid access token, prompting login if needed."""
        accounts = self.app.get_accounts()
        
        if accounts:
            # Try silent token acquisition (refresh token)
            result = self.app.acquire_token_silent(SCOPES, account=accounts[0])
            if result and "access_token" in result:
                self._save_cache()
                return result["access_token"]
        
        # Interactive login required
        result = self.app.acquire_token_interactive(
            SCOPES,
            port=8400,
            prompt="select_account",
        )
        
        if "access_token" in result:
            self._save_cache()
            return result["access_token"]
        
        raise Exception("Login failed: {}".format(result.get("error_description", "Unknown")))
    
    def get_accounts(self):
        """Returns cached MSAL accounts."""
        return self.app.get_accounts()
    
    def logout(self):
        """Clears token cache."""
        for acct in self.app.get_accounts():
            self.app.remove_account(acct)
        self._save_cache()
    
    def _save_cache(self):
        if self.cache.has_state_changed:
            os.makedirs(os.path.dirname(self._cache_path), exist_ok=True)
            with open(self._cache_path, "w") as f:
                f.write(self.cache.serialize())
```

**Key decisions:**
- Uses **PublicClientApplication** (no client secret needed — safe for desktop apps)
- Token cache persisted to `%LOCALAPPDATA%\InboxBar\graph_token_cache.bin`
- **Silent refresh** tried first; interactive login only when needed
- Port 8400 for localhost redirect URI

---

### Phase 3: Graph API Client Implementation (Day 2, ~6 hours)

#### 3.5 Core Graph API Client
**File: `sidebar/services/graph_client.py`**

Each method maps to a Microsoft Graph REST endpoint:

| OutlookClient Method | Graph API Endpoint | HTTP |
|---|---|---|
| `get_accounts()` | `/me` | GET |
| `get_inbox_items()` | `/me/mailFolders/inbox/messages` | GET |
| `get_unread_count()` | `/me/mailFolders/inbox` (unreadItemCount) | GET |
| `mark_as_read()` | `/me/messages/{id}` | PATCH |
| `delete_email()` | `/me/messages/{id}` | DELETE |
| `toggle_flag()` | `/me/messages/{id}` (flag property) | PATCH |
| `unflag_email()` | `/me/messages/{id}` (flag property) | PATCH |
| `get_calendar_items()` | `/me/calendarView` | GET |
| `get_tasks()` | `/me/todo/lists/{id}/tasks` | GET |
| `mark_task_complete()` | `/me/todo/lists/{id}/tasks/{id}` | PATCH |
| `create_email()` | Opens `mailto:` URL or POST `/me/messages` | POST |
| `create_meeting()` | POST `/me/events` | POST |
| `create_task()` | POST `/me/todo/lists/{id}/tasks` | POST |
| `check_new_mail()` | `/me/mailFolders/inbox/messages?$top=1&$orderby=receivedDateTime desc` | GET |
| `get_category_map()` | `/me/outlook/masterCategories` | GET |
| `search_contacts()` | `/me/contacts` + `/me/people` | GET |
| `get_folder_list()` | `/me/mailFolders` (recursive) | GET |
| `get_pulse_status()` | Combines calendar + tasks checks | GET |

#### Data Format Mapping
The Graph API returns JSON. We need to map it to match the dict format the UI expects:

```python
# Graph API message → InboxBar email dict
def _map_message(self, msg):
    return {
        "entry_id": msg["id"],           # Graph uses string IDs
        "store_id": None,                 # Not applicable for Graph
        "subject": msg.get("subject", ""),
        "sender": msg.get("from", {}).get("emailAddress", {}).get("name", ""),
        "sender_email": msg.get("from", {}).get("emailAddress", {}).get("address", ""),
        "received": datetime.fromisoformat(msg["receivedDateTime"].rstrip("Z")),
        "unread": not msg.get("isRead", True),
        "has_attachment": msg.get("hasAttachments", False),
        "importance": msg.get("importance", "normal"),
        "flag_status": 1 if msg.get("flag", {}).get("flagStatus") == "flagged" else 0,
        "flag_request": msg.get("flag", {}).get("flagStatus", ""),
        "flag_due": msg.get("flag", {}).get("dueDateTime", {}).get("dateTime"),
        "categories": msg.get("categories", []),
        "body_preview": msg.get("bodyPreview", ""),
        "conversation_id": msg.get("conversationId", ""),
    }

# Graph API event → InboxBar meeting dict
def _map_event(self, evt):
    return {
        "entry_id": evt["id"],
        "subject": evt.get("subject", ""),
        "start": datetime.fromisoformat(evt["start"]["dateTime"]),
        "end": datetime.fromisoformat(evt["end"]["dateTime"]),
        "location": evt.get("location", {}).get("displayName", ""),
        "response_status": self._map_response_status(evt.get("responseStatus", {}).get("response")),
        "is_recurring": evt.get("recurrence") is not None,
        "organizer": evt.get("organizer", {}).get("emailAddress", {}).get("name", ""),
    }
```

#### Key Differences from COM:

| Feature | COM (Current) | Graph API |
|---|---|---|
| **IDs** | Binary EntryID (hex string) | String GUID-like IDs |
| **Opening items** | `item.Display()` (opens in Outlook) | Open via `webLink` URL in browser |
| **Creating items** | `CreateItem()` opens in Outlook | Either POST to API or open `mailto:` / web URL |
| **Folder subfolder path** | Navigate via COM folders | `/me/mailFolders/{id}/childFolders` |
| **Offline** | Works | Doesn't work |
| **Speed** | Instant (local) | Network latency (~100-300ms per call) |
| **Batch operations** | Sequential COM calls | Can use `$batch` endpoint |

---

### Phase 4: UI Integration (Day 2, ~2 hours)

#### 3.6 Login/Account UI
**New: Settings panel additions**

- **"Backend" dropdown**: Auto / Classic Outlook / Microsoft 365
- **"Sign in to Microsoft 365" button**: Triggers OAuth flow
- **Account status indicator**: Shows logged-in account name
- **"Sign out" button**: Clears tokens

#### 3.7 "Open" Behaviour Change
- **COM**: `item.Display()` — opens in desktop Outlook
- **Graph**: `webbrowser.open(msg["webLink"])` — opens in browser/New Outlook

#### 3.8 "Create" Behaviour Change  
- **COM**: `outlook.CreateItem(0)` — opens compose window in Outlook
- **Graph**: `webbrowser.open("https://outlook.office.com/mail/deeplink/compose")` or POST to API

---

### Phase 5: Sidebar_main.py Refactoring (Day 2, ~2 hours)

#### 3.9 Replace Direct COM References
There are **~33 places** in `sidebar_main.py` that reference `self.outlook_client`. Most work through the interface, but a few access COM internals directly:

| Line | Current Code | Issue | Fix |
|---|---|---|---|
| 681 | `self.outlook_client.outlook` | Accesses raw COM object | Add `get_native_app()` to interface |
| 682 | `self.outlook_client.outlook` | Same | Same |
| 995 | `self.outlook_client.namespace.GetItemFromID()` | COM-specific | Use `get_item_by_entryid()` |
| 3027 | `self.outlook_client.outlook.CreateItem(0)` | COM-specific | Use `create_email()` |
| 3029 | `self.outlook_client.outlook.CreateItem(1)` | COM-specific | Use `create_meeting()` |
| 3034 | `self.outlook_client.outlook.CreateItem(3)` | COM-specific | Use `create_task()` |

These 6 direct COM references need to be routed through the abstract interface.

---

### Phase 6: PyInstaller & Build (Day 3, ~1 hour)

#### 3.10 Update InboxBar.spec
```python
hiddenimports=[
    # ... existing ...
    'msal',
    'requests',
    'requests.adapters',
    'urllib3',
    'certifi',
],
datas=[
    # ... existing ...
    # MSAL may need its metadata
],
```

#### 3.11 Update setup.iss (Installer)
No changes needed — new Python files are bundled automatically by PyInstaller.

---

### Phase 7: Testing (Day 3, ~3 hours)

#### 3.12 Test Matrix
| Test | COM | Graph |
|---|---|---|
| Launch + auto-detect backend | ✅ | ✅ |
| Sign in (OAuth flow) | N/A | ✅ |
| List unread emails | ✅ | ✅ |
| Mark as read | ✅ | ✅ |
| Delete email | ✅ | ✅ |
| Flag/unflag email | ✅ | ✅ |
| Calendar items | ✅ | ✅ |
| Tasks | ✅ | ✅ |
| Complete task | ✅ | ✅ |
| Quick create (email, meeting, task) | ✅ | ✅ |
| New mail detection + pulse | ✅ | ✅ |
| Categories + colors | ✅ | ✅ |
| Search contacts | ✅ | ✅ |
| Open item | ✅ Desktop Outlook | ✅ Browser |
| Sharing/send email | ✅ | ✅ |
| Token refresh (silent) | N/A | ✅ |
| Token expiry + re-login | N/A | ✅ |

---

## 4. File Summary

### New Files
| File | Purpose |
|---|---|
| `sidebar/services/mail_client.py` | Abstract interface (ABC) |
| `sidebar/services/graph_auth.py` | MSAL token management & OAuth flow |
| `sidebar/services/graph_client.py` | Graph API implementation of MailClient |

### Modified Files
| File | Changes |
|---|---|
| `sidebar/services/outlook_client.py` | Inherit from MailClient ABC |
| `sidebar/core/config_manager.py` | Add `backend` setting |
| `sidebar_main.py` | Backend selection, remove 6 direct COM refs |
| `sidebar/ui/panels/settings.py` | Backend selector + sign-in button |
| `InboxBar.spec` | Add msal/requests to hidden imports |
| `requirements.txt` | Add msal, requests |

### Unchanged
- All UI rendering code (email cards, calendar, tasks, etc.)
- Theme system
- Toolbar, help panel, share dialog
- Update checker
- Pulse animation

---

## 5. Timeline Estimate

| Phase | Effort | Description |
|---|---|---|
| **Phase 1** | 2 hours | Abstract interface, backend selection |
| **Phase 2** | 3 hours | MSAL auth, token cache, OAuth flow |
| **Phase 3** | 6 hours | Full Graph API client (~20 methods) |
| **Phase 4** | 2 hours | Settings UI for login/backend |
| **Phase 5** | 2 hours | Refactor sidebar_main.py |
| **Phase 6** | 1 hour | Build & packaging updates |
| **Phase 7** | 3 hours | End-to-end testing |
| **Total** | **~19 hours** | ~2-3 working days |

---

## 6. Risks & Mitigations

| Risk | Likelihood | Mitigation |
|---|---|---|
| Azure AD admin consent required | Medium | Use delegated permissions (user-level), no admin needed |
| Token expires during use | Low | MSAL handles refresh automatically |
| Graph API rate limits | Low | InboxBar polls every 30s, well under limits |
| Different data formats break UI | Medium | Thorough data mapping layer |
| "Open" behaviour feels different | Medium | Clear UX — explain it opens in browser |
| Network latency vs COM speed | High | Show loading indicators, cache results |
| Users confused by OAuth popup | Medium | Clear first-run instructions in UI |

---

## 7. Decision Points (Need Your Input)

1. **Azure AD Registration**: Do you have access to register apps in your M365 tenant, or should we target personal Microsoft accounts?

2. **"Open" in browser acceptable?**: With Graph, clicking "Open" on an email opens it in a browser tab (Outlook web). Is that OK, or do you want a preview pane?

3. **Priority**: Should we keep COM as the primary backend (auto-detect) and Graph as fallback, or the reverse?

4. **Quick Create**: With Graph, creating a new email/meeting can either:
   - Open a compose URL in the browser (simpler)
   - POST via API and show a custom compose UI in the sidebar (complex)
   
   Which do you prefer?

5. **Tasks API**: Microsoft has two task systems:
   - **To Do** (modern, via Graph `/me/todo/lists/.../tasks`)
   - **Outlook Tasks** (legacy, limited Graph support)
   
   Your current COM client uses Outlook Tasks. Should we migrate to To Do, support both, or just To Do?
```
