from sidebar.services.mail_client import MailClient
from sidebar.services.outlook_client import OutlookClient
from sidebar.services.graph_client import GraphAPIClient

class HybridMailClient(MailClient):
    """
    Multiplexes both Classic Outlook (COM) and Microsoft 365 (Graph API).
    Allows user to have COM accounts and Graph accounts running simultaneously.
    """
    def __init__(self):
        self.com = OutlookClient()
        self.graph = GraphAPIClient()
        self.last_received_time = None

    def _split_accounts(self, account_names):
        """Splits the requested account names into COM and Graph based on known accounts."""
        if not account_names:
            # If none specified, query both
            return self.com.get_accounts(), self.graph.get_accounts()
            
        com_known = self.com.get_accounts()
        graph_known = self.graph.get_accounts()
        
        c_accs = [a for a in account_names if a in com_known]
        g_accs = [a for a in account_names if a in graph_known or a not in com_known] # Default unknown to Graph? Or just skip
        
        return c_accs, g_accs

    def connect(self) -> bool:
        c = self.com.connect()
        g = self.graph.connect()
        return c or g

    def reconnect(self) -> bool:
        c = self.com.reconnect()
        g = self.graph.reconnect()
        return c or g

    def is_connected(self) -> bool:
        return self.com.is_connected() or self.graph.is_connected()

    def get_accounts(self) -> list:
        # Combine unique accounts from both
        accs = []
        for a in self.com.get_accounts():
            if a not in accs: accs.append(a)
        for a in self.graph.get_accounts():
            if a not in accs: accs.append(a)
        return accs

    def get_inbox_items(self, count=20, unread_only=False, only_flagged=False, 
                        due_filters=None, account_names=None, account_config=None) -> tuple:
        c_names, g_names = self._split_accounts(account_names)
        
        all_emails = []
        total_unread = 0
        
        if c_names or not account_names:
            c_emails, c_unread = self.com.get_inbox_items(count, unread_only, only_flagged, due_filters, c_names, account_config)
            all_emails.extend(c_emails)
            total_unread += c_unread
            
        if g_names or not account_names:
            g_emails, g_unread = self.graph.get_inbox_items(count, unread_only, only_flagged, due_filters, g_names, account_config)
            all_emails.extend(g_emails)
            total_unread += g_unread
            
        # Sort combined by received date (safely handling None dates)
        import datetime
        min_date = datetime.datetime.min
        all_emails.sort(key=lambda x: x.get("received") or min_date, reverse=True)
        return all_emails[:count], total_unread

    def get_unread_count(self, account_names=None, account_config=None) -> int:
        c_names, g_names = self._split_accounts(account_names)
        total = 0
        if c_names or not account_names: total += self.com.get_unread_count(c_names, account_config)
        if g_names or not account_names: total += self.graph.get_unread_count(g_names, account_config)
        return total

    def _route_item(self, entry_id, store_id, method_name, *args):
        # We don't strictly know which backend owns an entry_id unless we check both or assume based on ID format
        # Graph IDs are very long Base64 looking strings. COM EntryIDs are hex.
        # Simple heuristic: if it looks like hex, try COM. Else try Graph.
        # But safest is trying COM, if fails, try Graph.
        import string
        is_hex = all(c in string.hexdigits for c in entry_id) if hasattr(entry_id, 'isalnum') else False
        
        if is_hex and len(entry_id) > 40:
             # Looks like COM EntryID
             try:
                 return getattr(self.com, method_name)(entry_id, store_id, *args)
             except: pass
        
        # Try Graph
        try:
             func = getattr(self.graph, method_name)
             return func(entry_id, store_id, *args)
        except Exception as e:
             print(f"DEBUG: Graph exception on {method_name}: {e}")
             import traceback
             traceback.print_exc()
             return None

    def mark_as_read(self, entry_id, store_id=None) -> bool:
        return self._route_item(entry_id, store_id, "mark_as_read")

    def delete_email(self, entry_id, store_id=None) -> bool:
        return self._route_item(entry_id, store_id, "delete_email")

    def toggle_flag(self, entry_id, store_id=None) -> bool:
        return self._route_item(entry_id, store_id, "toggle_flag")

    def unflag_email(self, entry_id, store_id=None) -> bool:
        return self._route_item(entry_id, store_id, "unflag_email")

    def open_item(self, entry_id, store_id=None):
        return self._route_item(entry_id, store_id, "open_item")

    def reply_to_email(self, entry_id, store_id=None):
        return self._route_item(entry_id, store_id, "reply_to_email")
        
    def move_email(self, entry_id, folder_name, store_id=None):
        # We need to explicitly pass folder_name as an arg to the router
        return self._route_item(entry_id, store_id, "move_email", folder_name)

    def get_item_by_entryid(self, entry_id, store_id=None):
        return self._route_item(entry_id, store_id, "get_item_by_entryid")

    def get_calendar_items(self, start_dt, end_dt, account_names=None) -> list:
        c_names, g_names = self._split_accounts(account_names)
        items = []
        if c_names: items.extend(self.com.get_calendar_items(start_dt, end_dt, c_names))
        if g_names: items.extend(self.graph.get_calendar_items(start_dt, end_dt, g_names))
        # Standardize timezone before sorting (make naive)
        for item in items:
            dt = item.get("start")
            if dt and dt.tzinfo is not None:
                item["start"] = dt.replace(tzinfo=None)
                
        # Sort by start time
        items.sort(key=lambda x: x.get("start"))
        return items

    def get_tasks(self, due_filters=None, account_names=None) -> list:
        c_names, g_names = self._split_accounts(account_names)
        tasks = []
        if c_names: tasks.extend(self.com.get_tasks(due_filters, c_names))
        if g_names: tasks.extend(self.graph.get_tasks(due_filters, g_names))
        
        def sort_key(x):
            try:
                if x.get('has_reminder'): return (0, x.get('due') or 0)
                if x.get('importance') == 'High': return (1, x.get('due') or 0)
                return (2, x.get('due') or 0)
            except:
                return (3, 0)
        
        tasks.sort(key=sort_key)
        return tasks

    def mark_task_complete(self, entry_id, store_id=None) -> bool:
        return self._route_item(entry_id, store_id, "mark_task_complete")

    def create_email(self):
        # Default to COM if available, else Graph
        if self.com.is_connected(): self.com.create_email()
        elif self.graph.is_connected(): self.graph.create_email()

    def create_meeting(self):
        if self.com.is_connected(): self.com.create_meeting()
        elif self.graph.is_connected(): self.graph.create_meeting()

    def create_task(self):
        if self.com.is_connected(): self.com.create_task()
        elif self.graph.is_connected(): self.graph.create_task()

    def create_contact(self):
        if self.com.is_connected(): self.com.create_contact()
        elif self.graph.is_connected(): self.graph.create_contact()

    def check_new_mail(self, account_names=None) -> bool:
        c_names, g_names = self._split_accounts(account_names)
        val = False
        if c_names: val = val or self.com.check_new_mail(c_names)
        if g_names: val = val or self.graph.check_new_mail(g_names)
        return val

    def get_pulse_status(self, account_names=None) -> dict:
        c_names, g_names = self._split_accounts(account_names)
        p1 = self.com.get_pulse_status(c_names) if c_names else {"calendar": None, "tasks": None}
        p2 = self.graph.get_pulse_status(g_names) if g_names else {"calendar": None, "tasks": None}
        
        # Combine (take highest urgency)
        res = {"calendar": p1.get("calendar") or p2.get("calendar"), 
               "tasks": p1.get("tasks") or p2.get("tasks")}
        return res

    def get_category_map(self) -> dict:
        # Merge dicts
        m = self.com.get_category_map()
        m.update(self.graph.get_category_map())
        return m

    def search_contacts(self, query, max_results=8) -> list:
        # Only search COM or Graph?
        res = self.com.search_contacts(query, max_results)
        if not res:
            res = self.graph.search_contacts(query, max_results)
        return res

    def get_folder_list(self, account_name=None) -> list:
        # Route to correct client
        if account_name in self.com.get_accounts():
            return self.com.get_folder_list(account_name)
        return self.graph.get_folder_list(account_name)

    def get_native_app(self):
        return self.com.get_native_app()

    def send_email_with_attachment(self, recipient, subject, body, attachment_path) -> bool:
        if self.com.is_connected():
            return self.com.send_email_with_attachment(recipient, subject, body, attachment_path)
        return False
