import requests
import webbrowser
from datetime import datetime, timedelta
import urllib.parse
from urllib.parse import quote
from sidebar.services.mail_client import MailClient
from sidebar.services.graph_auth import GraphAuth
import traceback

class GraphAPIClient(MailClient):
    """
    Microsoft Graph API (HTTP) implementation of the MailClient interface.
    Used for New Outlook, Office 365, and web-only users.
    """
    def __init__(self):
        self.auth = GraphAuth()
        self.base_url = "https://graph.microsoft.com/v1.0"
        self._cache = {}
        self.last_received_time = None
        self._connected = False

    # --- Core HTTP Helper ---
    def _request(self, method, endpoint, **kwargs):
        """Helper to append token and execute a Graph API request."""
        # Try to get a valid token without prompting
        token = self.auth.get_token(interactive=False)
        if not token:
            # Silently return if not authenticated to prevent log spam in hybrid mode
            return None

        headers = kwargs.pop('headers', {})
        headers['Authorization'] = f"Bearer {token}"
        headers['Content-Type'] = 'application/json'
        
        url = endpoint if endpoint.startswith("http") else f"{self.base_url}{endpoint}"
        
        # Enforce a connection timeout so the app doesn't hang when offline.
        # (connect_timeout=10s, read_timeout=30s)
        if 'timeout' not in kwargs:
            kwargs['timeout'] = (10, 30)
        
        try:
            resp = requests.request(method, url, headers=headers, **kwargs)
            if resp.status_code == 204: # No Content (Success)
                return True
            resp.raise_for_status()
            
            # Not all APIs return JSON (e.g., photo downloads)
            if "application/json" in resp.headers.get("Content-Type", ""):
                 return resp.json()
            return resp.content
        except (requests.exceptions.ConnectionError, requests.exceptions.Timeout) as e:
            # Network/connectivity errors — re-raise so callers can detect offline state
            print(f"[Graph] Network error {method} {endpoint}: {e}")
            raise
        except requests.exceptions.RequestException as e:
            # HTTP errors (4xx/5xx) — return None, these are API-level issues
            print(f"[Graph] API Error {method} {endpoint}: {e}")
            if hasattr(e, 'response') and e.response is not None and hasattr(e.response, 'text'):
                print(f"[Graph] detail: {e.response.text}")
            return None

    # --- Connection & Auth ---
    def _get_domain(self):
        email = self.auth.get_current_user_email()
        if not email: return "office.com"
        consumer_domains = ["@outlook.com", "@hotmail.com", "@live.com", "@msn.com"]
        if any(email.lower().endswith(d) for d in consumer_domains):
            return "live.com"
        return "office.com"

    def connect(self) -> bool:
        """Prompt user to sign in if not already signed in."""
        try:
            token = self.auth.get_token(interactive=True)
            self._connected = token is not None
            return self._connected
        except Exception as e:
            print(f"[Graph] Connect error: {e}")
            return False

    def reconnect(self) -> bool:
        # Just attempts to get token silently
        token = self.auth.get_token(interactive=False)
        self._connected = token is not None
        return self._connected

    def is_connected(self) -> bool:
        return self._connected

    def get_accounts(self) -> list:
        # Return the currently logged in account
        email = self.auth.get_current_user_email()
        return [email] if email else []

    # --- Data Mappers ---
    def _map_message(self, msg):
        """Converts Graph API Message JSON to InboxBar schema."""
        if not msg: return None
        
        try:
            # 2024-03-24T15:30:00Z -> standard datetime
            recv_str = msg.get("receivedDateTime", "")
            if recv_str.endswith("Z"):
                recv_dt = datetime.fromisoformat(recv_str[:-1])
            else:
                 # fallback format?
                recv_dt = datetime.strptime(recv_str[:19], "%Y-%m-%dT%H:%M:%S")
        except:
             recv_dt = datetime.now()
             
        # Flag parsing
        flag_status = 0 # olNoFlag
        flagObj = msg.get("flag", {})
        fStatus = flagObj.get("flagStatus", "unflagged")
        fReq = ""
        fDue = None
        
        if fStatus == "flagged":
            flag_status = 1 # olFlagMarked
        elif fStatus == "complete":
            flag_status = 2 # olFlagComplete

        from_email = msg.get("from", {}).get("emailAddress", {})
        
        return {
            "entry_id": msg.get("id"),
            "store_id": None, # Graph doesn't use store IDs
            "subject": msg.get("subject", ""),
            "sender": from_email.get("name", "Unknown"),
            "sender_email": from_email.get("address", ""),
            "received": recv_dt,
            "unread": not msg.get("isRead", True),
            "has_attachment": msg.get("hasAttachments", False),
            "importance": msg.get("importance", "normal"),
            "flag_status": flag_status,
            "flag_request": fReq, # API doesn't return 'Follow up' string natively?
            "flag_due": fDue,
            "categories": msg.get("categories", []),
            "body_preview": msg.get("bodyPreview", ""),
            "conversation_id": msg.get("conversationId", ""),
            "web_link": msg.get("webLink", ""),  # Specific to Graph
        }

    def _map_event(self, evt):
         """Converts Graph API Event JSON to InboxBar schema."""
         if not evt: return None
         
         # Time handling (usually in UTC for Graph API /me/calendarView if timezone header not set)
         start_str = evt.get("start", {}).get("dateTime", "")
         end_str = evt.get("end", {}).get("dateTime", "")
         
         try:
            start_dt = datetime.fromisoformat(start_str.split(".")[0])
            end_dt = datetime.fromisoformat(end_str.split(".")[0])
         except:
             start_dt = end_dt = datetime.now()
         
         # Response status
         rsp_map = {
             "none": 0,
             "organizer": 1,
             "tentativelyAccepted": 2,
             "accepted": 3,
             "declined": 4,
             "notResponded": 5
         }
         r_str = evt.get("responseStatus", {}).get("response", "none")
         rsp = rsp_map.get(r_str, 0)
         
         link = evt.get("webLink", "")
         email = self.auth.get_current_user_email()
         if email and link:
             link += f"&login_hint={urllib.parse.quote(email)}" if "?" in link else f"?login_hint={urllib.parse.quote(email)}"
             
         return {
            "entry_id": evt.get("id"),
            "subject": evt.get("subject", "No Title"),
            "start": start_dt,
            "end": end_dt,
            "location": evt.get("location", {}).get("displayName", ""),
            "response_status": rsp,
            "organizer": evt.get("organizer", {}).get("emailAddress", {}).get("name", ""),
            "is_recurring": evt.get("seriesMasterId") is not None,
            "web_link": link,
         }

    def _map_task(self, task):
        """Converts Graph API To Do Task JSON to InboxBar schema."""
        if not task: return None
        
        due_dt = None
        due_obj = task.get("dueDateTime")
        if due_obj and "dateTime" in due_obj:
            try:
                 due_dt = datetime.fromisoformat(due_obj["dateTime"].split(".")[0])
            except: pass
            
        email = self.auth.get_current_user_email()
        login_hint = f"?login_hint={urllib.parse.quote(email)}" if email else ""
        return {
            "entry_id": task.get("id"),
            "subject": task.get("title", ""),
            "due": due_dt,
            "importance": task.get("importance", "normal").title(),
            "status": "Completed" if task.get("status") == "completed" else "NotStarted",
            "categories": task.get("categories", []), # Often empty for ToDo unless mapped
            "is_recurring": False, # Basic task structure
            "has_reminder": task.get("isReminderOn", False),
            "complete": task.get("status") == "completed",
            "web_link": f"https://to-do.{self._get_domain()}/tasks/id/{task.get('id')}{login_hint}",
        }

    # --- Email Operations ---
    def get_inbox_items(self, count=20, unread_only=False, only_flagged=False, 
                        due_filters=None, account_names=None, account_config=None) -> tuple:
        """Fetch emails via /me/mailFolders/inbox/messages"""
        
        filters = []
        if unread_only:
            filters.append("isRead eq false")
        # Removing `flag/flagStatus eq 'flagged'` because Microsoft Graph blocks this as an "InefficientFilter" on Inbox
            
        params = [
            f"$top={count}",
            "$orderby=receivedDateTime desc"
        ]
        if filters:
             params.append(f"$filter=" + " and ".join(filters))
             
        query = "&".join(params)
        endpoint = f"/me/mailFolders/inbox/messages?{query}"
        
        data = self._request("GET", endpoint)
        if not data:
            return [], 0
            
        messages = [self._map_message(m) for m in data.get("value", [])]
        
        # Apply flag filter locally since Graph API rejects it conditionally
        if only_flagged:
             messages = [m for m in messages if m["flag_status"] != 0]

        # We also need total unread count (Graph doesn't return total DB count on a filtered query)
        unread_count = self.get_unread_count()
        
        # Track latest time
        if messages:
            latest = messages[0]["received"]
            if not self.last_received_time or latest > self.last_received_time:
                self.last_received_time = latest
                
        return messages, unread_count

    def get_unread_count(self, account_names=None, account_config=None) -> int:
        data = self._request("GET", "/me/mailFolders/inbox?$select=unreadItemCount")
        if data:
            return data.get("unreadItemCount", 0)
        return 0

    def mark_as_read(self, entry_id, store_id=None) -> bool:
        resp = self._request("PATCH", f"/me/messages/{entry_id}", json={"isRead": True})
        return resp is not None

    def delete_email(self, entry_id, store_id=None) -> bool:
        resp = self._request("DELETE", f"/me/messages/{entry_id}")
        return resp is not None

    def toggle_flag(self, entry_id, store_id=None) -> bool:
        # 1. Fetch current
        msg = self._request("GET", f"/me/messages/{entry_id}?$select=flag")
        if not msg: return False
        
        status = msg.get("flag", {}).get("flagStatus", "unflagged")
        new_status = "notFlagged" if status == "flagged" else "flagged"
        
        resp = self._request("PATCH", f"/me/messages/{entry_id}", json={"flag": {"flagStatus": new_status}})
        return resp is not None

    def unflag_email(self, entry_id, store_id=None) -> bool:
        resp = self._request("PATCH", f"/me/messages/{entry_id}", json={"flag": {"flagStatus": "notFlagged"}})
        return resp is not None

    def open_item(self, entry_id, store_id=None):
        """Open web link in default browser."""
        msg = self._request("GET", f"/me/messages/{entry_id}?$select=webLink")
        if msg and "webLink" in msg:
            link = msg["webLink"]
            email = self.auth.get_current_user_email()
            if email and link:
                link += f"&login_hint={urllib.parse.quote(email)}" if "?" in link else f"?login_hint={urllib.parse.quote(email)}"
            webbrowser.open(link)
        else:
            print("[Graph] Could not find webLink to open item.")

    def reply_to_email(self, entry_id, store_id=None) -> bool:
        """Opens a compose window directly from Graph API via deeplink or creates a draft."""
        # The most reliable way cross-platform without creating headless drafts: Let the browser do it if possible,
        # otherwise create a reply draft and open it.
        # Create draft:
        resp = self._request("POST", f"/me/messages/{entry_id}/createReply")
        if resp and "id" in resp:
            # Re-fetch for link
            msg = self._request("GET", f"/me/messages/{resp['id']}?$select=webLink")
            if msg and "webLink" in msg:
                link = msg["webLink"]
                email = self.auth.get_current_user_email()
                if email and link:
                    link += f"&login_hint={urllib.parse.quote(email)}" if "?" in link else f"?login_hint={urllib.parse.quote(email)}"
                webbrowser.open(link)
                return True
        return False

    def move_email(self, entry_id, folder_name, store_id=None) -> bool:
        """Moves an email to a destination folder by name."""
        # Find folder ID
        folders = self._request("GET", f"/me/mailFolders?$filter=displayName eq '{folder_name}'")
        if not folders or not folders.get("value"):
            print(f"[Graph] Could not find folder '{folder_name}'")
            return False
            
        target_id = folders["value"][0]["id"]
        
        # Execute Move
        resp = self._request("POST", f"/me/messages/{entry_id}/move", json={"destinationId": target_id})
        return resp is not None

    # --- Calendar ---
    def get_calendar_items(self, start_dt, end_dt, account_names=None) -> list:
        # Add timezone header so returned times match local
        s_iso = start_dt.isoformat()
        e_iso = end_dt.isoformat()
        
        endpoint = f"/me/calendarView?startDateTime={s_iso}&endDateTime={e_iso}&$top=50&$orderby=start/dateTime"
        
        # Preferred timezone is essential here to align with desktop
        import time
        tz_offset = -time.timezone // 3600
        tz_name = f"{'UTC' if tz_offset == 0 else 'Etc/GMT'}{'+' if tz_offset <= 0 else '-'}{abs(tz_offset)}"
        
        headers = {"Prefer": 'outlook.timezone="UTC"'} # Default to UTC, process locallly
        # For simplicity, returning UTC and mapping locally
        data = self._request("GET", endpoint, headers=headers)
        
        if not data: return []
        return [self._map_event(e) for e in data.get("value", []) if e.get("isCancelled") != True]

    # --- Tasks (To Do Lists) ---
    def get_tasks(self, due_filters=None, account_names=None) -> list:
        # First, ensure we have the default To Do list ID
        if "todo_list_id" not in self._cache:
             lists = self._request("GET", "/me/todo/lists")
             if lists and "value" in lists and lists["value"]:
                 self._cache["todo_list_id"] = lists["value"][0]["id"]
             else:
                 return []
                 
        list_id = self._cache["todo_list_id"]
        
        # Get tasks that are NOT completed
        endpoint = f"/me/todo/lists/{list_id}/tasks?$filter=status ne 'completed'&$top=50"
        data = self._request("GET", endpoint)
        
        if not data: return []
        
        tasks = [self._map_task(t) for t in data.get("value", [])]
        
        # Apply client-side filtering for Due Dates
        # Similar logic to COM
        filtered_tasks = []
        now_dt = datetime.now()
        now_date = now_dt.date()
        
        if not due_filters: due_filters = ["Overdue", "Today", "Tomorrow"]
        
        for task in tasks:
            if not task["due"]:
                if "No Date" in due_filters:
                    filtered_tasks.append(task)
                continue
                
            task_date = task["due"].date()
            if task_date < now_date and "Overdue" in due_filters:
                filtered_tasks.append(task)
            elif task_date == now_date and "Today" in due_filters:
                filtered_tasks.append(task)
            elif task_date == now_date + timedelta(days=1) and "Tomorrow" in due_filters:
                filtered_tasks.append(task)
                
        return filtered_tasks

    def mark_task_complete(self, entry_id, store_id=None) -> bool:
         # Need list_id
         if "todo_list_id" not in self._cache: return False
         list_id = self._cache["todo_list_id"]
         
         resp = self._request("PATCH", f"/me/todo/lists/{list_id}/tasks/{entry_id}", json={"status": "completed"})
         return resp is not None

    # --- Quick Create ---
    def create_email(self):
        # Open Outlook Web Compose
        email = self.auth.get_current_user_email()
        login_hint = f"?login_hint={urllib.parse.quote(email)}" if email else ""
        webbrowser.open(f"https://outlook.{self._get_domain()}/mail/deeplink/compose{login_hint}")

    def create_meeting(self):
        email = self.auth.get_current_user_email()
        login_hint = f"?login_hint={urllib.parse.quote(email)}" if email else ""
        webbrowser.open(f"https://outlook.{self._get_domain()}/calendar/deeplink/compose{login_hint}")

    def create_task(self):
        email = self.auth.get_current_user_email()
        login_hint = f"?login_hint={urllib.parse.quote(email)}" if email else ""
        webbrowser.open(f"https://to-do.{self._get_domain()}/tasks/today{login_hint}")
        
    def create_contact(self):
         email = self.auth.get_current_user_email()
         login_hint = f"?login_hint={urllib.parse.quote(email)}" if email else ""
         webbrowser.open(f"https://outlook.{self._get_domain()}/people/{login_hint}")

    # --- General / Utility ---
    def check_new_mail(self, account_names=None) -> bool:
        """Returns True if a new mail has arrived since last poll."""
        data = self._request("GET", "/me/mailFolders/inbox/messages?$top=1&$orderby=receivedDateTime desc&$select=receivedDateTime")
        if data and "value" in data and len(data["value"]) > 0:
            latest_str = data["value"][0]["receivedDateTime"]
            try:
                latest_dt = datetime.fromisoformat(latest_str.rstrip("Z"))
                
                # First run or cache
                if not self.last_received_time:
                    self.last_received_time = latest_dt
                    return False
                    
                if latest_dt > self.last_received_time:
                    self.last_received_time = latest_dt
                    return True
            except: pass
        return False

    def get_pulse_status(self, account_names=None) -> dict:
        status = {"calendar": None, "tasks": None}
        
        # Not fully implemented yet due to rate limiting vs value
        # But we could check next event time here
        return status

    def get_category_map(self) -> dict:
        # Cache this as it doesn't change often
        if "categories" in self._cache: return self._cache["categories"]
        
        data = self._request("GET", "/me/outlook/masterCategories")
        cmap = {}
        if data and "value" in data:
            for cat in data["value"]:
                 # Map preset names to RGB hex or standard index
                 # This would need mapping logic. For now return empty.
                 cmap[cat.get("displayName")] = -1
        
        self._cache["categories"] = cmap
        return cmap

    def search_contacts(self, query, max_results=8) -> list:
        q = quote(query)
        data = self._request("GET", f"/me/people?$search=\"{q}\"&$top={max_results}")
        res = []
        if data and "value" in data:
            for p in data["value"]:
                res.append({
                    "name": p.get("displayName", ""),
                    "email": p.get("scoredEmailAddresses", [{}])[0].get("address", "")
                })
        return res

    def get_folder_list(self, account_name=None) -> list:
        # Just return top level for now. Full recursion takes multiple API calls.
        data = self._request("GET", "/me/mailFolders?$top=20")
        if not data: return ["Inbox", "Sent Items", "Deleted Items"]
        return [f.get("displayName") for f in data.get("value", [])]

    def get_native_app(self):
         return None # No native app available.

    def send_email_with_attachment(self, recipient, subject, body, attachment_path) -> bool:
         """Drafts and sends an email via Graph API. Used for Share."""
         # Not fully implemented yet for attachment base64 encoding.
         print("[Graph] Share via background email not yet fully implemented.")
         return False
