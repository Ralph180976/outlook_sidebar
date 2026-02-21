# -*- coding: utf-8 -*-
import win32com.client
from datetime import datetime, timedelta
import pythoncom
import os

# Import theme constants from core
from sidebar.core.theme import OL_CAT_COLORS

class OutlookClient:
    # Re-expose for compatibility if needed, or just use the imported one
    OL_CAT_COLORS = OL_CAT_COLORS

    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.last_received_time = None
        self.connect()
        # Initialize last_received_time
        if self.namespace:
            self.check_latest_time()

    def connect(self):
        """Attempts to connect to the Outlook COM object."""
        try:
            # Helper to initialize COM in this thread if needed
            # pythoncom.CoInitialize() 
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            # print("Connected to Outlook")
            return True
        except Exception as e:
            print("Error connecting to Outlook: {}".format(e))
            self.outlook = None
            self.namespace = None
            return False

    def reconnect(self):
        """Force a full COM reconnection (e.g. after network change)."""
        print("COM reconnect: forcing full reconnection...")
        self.outlook = None
        self.namespace = None
        success = self.connect()
        if success:
            self.check_latest_time()
            print("COM reconnect: success")
        else:
            print("COM reconnect: failed")
        return success

    def _is_connection_healthy(self):
        """Quick probe to check if COM connection is still alive."""
        if not self.namespace:
            return False
        try:
            _ = self.namespace.Stores.Count
            return True
        except:
            return False

    def get_accounts(self):
        """Returns list of account names."""
        accounts = []
        if not self.connect(): return []
        try:
            for store in self.namespace.Stores:
                accounts.append(store.DisplayName)
        except Exception as e:
            print("Error fetching accounts: {}".format(e))
        return accounts

    def _get_enabled_stores(self, account_names):
        """Helper: Yields stores that match the provided names (or all if None)."""
        if not self.namespace: return
        
        try:
            for store in self.namespace.Stores:
                if account_names is None or store.DisplayName in account_names:
                    yield store
        except Exception:
            return

    def check_latest_time(self, account_names=None):
        """Updates the globally tracked last_received_time from enabled accounts using safe Tables."""
        if not self.namespace: return
        
        latest = None
        
        try:
            for store in self._get_enabled_stores(account_names):
                try:
                    inbox = store.GetDefaultFolder(6)
                    # Use Table to avoid traversing MailItem objects (Security Guard)
                    table = inbox.GetTable()
                    table.Sort("ReceivedTime", True) # Descending
                    table.Columns.RemoveAll()
                    table.Columns.Add("ReceivedTime")
                    
                    if not table.EndOfTable:
                        row = table.GetNextRow()
                        if row:
                            # Use GetValues for safety
                            vals = row.GetValues()
                            t = vals[0]
                            if latest is None or t > latest:
                                latest = t
                except:
                    continue
                    
            if latest:
                self.last_received_time = latest
                
        except Exception:
             pass

    def check_new_mail(self, account_names=None):
        """Checks for new mail across enabled accounts using safe Tables."""
        for attempt in range(2):
            if not self.namespace:
                if not self.connect(): return False

            # Detect stale COM connection (e.g. after network change)
            if not self._is_connection_healthy():
                print("COM connection stale in check_new_mail, reconnecting...")
                if not self.reconnect(): return False

            try:
                found_new = False
                global_max = self.last_received_time
                
                for store in self._get_enabled_stores(account_names):
                    try:
                        inbox = store.GetDefaultFolder(6)
                        # Use Table for security safety
                        table = inbox.GetTable()
                        table.Sort("ReceivedTime", True)
                        table.Columns.RemoveAll()
                        table.Columns.Add("ReceivedTime")
                        
                        if not table.EndOfTable:
                            row = table.GetNextRow()
                            if row:
                                vals = row.GetValues()
                                current_time = vals[0]
                                # Compare against our global max
                                if self.last_received_time and current_time > self.last_received_time:
                                    # print(f"DEBUG: Found NEW mail! {current_time} > {self.last_received_time}")
                                    found_new = True
                                
                                # Update local tracker for this poll
                                if global_max is None or current_time > global_max:
                                    global_max = current_time
                    except Exception as e:
                        # print(f"DEBUG: Error checking store: {e}")
                        continue
                
                if global_max:
                    self.last_received_time = global_max
                    
                return found_new
                
            except Exception as e:
                print("Polling error (Attempt {}): {}".format(attempt+1, e))
                self.namespace = None 
        
        return False

    def get_unread_count(self, account_names=None, account_config=None):
        """Returns total unread count from configured folders."""
        total = 0
        if not self.namespace: 
            if not self.connect(): return 0
            
        for store in self._get_enabled_stores(account_names):
            try:
                folders_to_scan = []
                # Check config
                if account_config and store.DisplayName in account_config:
                    conf = account_config[store.DisplayName]
                    if "email_folders" in conf and conf["email_folders"]:
                        for path in conf["email_folders"]:
                            f = self.get_folder_by_path(store, path)
                            if f: folders_to_scan.append(f)
                
                # Fallback to Inbox
                if not folders_to_scan:
                    try: folders_to_scan.append(store.GetDefaultFolder(6))
                    except: pass
                    
                for f in folders_to_scan:
                    try: total += f.UnReadItemCount
                    except: pass
            except: continue
        return total

    def get_folder_by_path(self, store, folder_path):
        """Helper to navigate folder path string (e.g. 'Inbox/Subfolder')."""
        try:
            parts = folder_path.split("/")
            curr = store.GetRootFolder()
            for p in parts:
                curr = curr.Folders[p]
            return curr
        except:
            return None

    def get_calendar_items(self, start_dt, end_dt, account_names=None):
        """Fetches calendar items from all enabled accounts. Accepts datetime objects."""
        # Print Debug Log
        try:
             # print("DEBUG: Calendar Query: {} to {}".format(start_dt, end_dt))
             pass
        except: pass

        for attempt in range(2):
            if not self.namespace:
                 if not self.connect(): return []
            try:
                all_results = []
                
                # Format for DASL/Jet - UK Format for this user's locale
                s_str = start_dt.strftime('%d/%m/%Y %H:%M')
                e_str = end_dt.strftime('%d/%m/%Y %H:%M')
                
                for store in self._get_enabled_stores(account_names):
                    try:
                        cal = store.GetDefaultFolder(9)
                        items = cal.Items
                        items.Sort("[Start]")
                        items.IncludeRecurrences = True
                        
                        restrict = "[Start] >= '{}' AND [Start] <= '{}'".format(s_str, e_str)
                        
                        try:
                            items = items.Restrict(restrict)
                        except Exception as e:
                            print("Restrict Warning: {}".format(e))
                            pass
                        
                        for item in items:
                            try:
                                # Manual Date Check (Safety Net against locale issues)
                                # Normalize Item Start (Aware/Naive)
                                i_start = item.Start
                                if getattr(i_start, "tzinfo", None) is not None:
                                     i_start = i_start.replace(tzinfo=None) # Make naive for comparison with our naive start_dt/end_dt
                                
                                if i_start < start_dt or i_start > end_dt:
                                     continue

                                all_results.append({
                                    "subject": item.Subject,
                                    "start": item.Start,
                                    "location": getattr(item, "Location", ""),
                                    "entry_id": item.EntryID,
                                    "is_meeting": True,
                                    "response_status": getattr(item, "ResponseStatus", 0),
                                    "account": store.DisplayName # Optional: Track source
                                })
                            except:
                                continue
                    except:
                        continue
                        
                # Sort merged results by start time
                try:
                    all_results.sort(key=lambda x: x["start"])
                except:
                    pass
                    
                return all_results
            except Exception as e:
                print("Calendar error: {}".format(e))
                self.namespace = None
        return []

    def get_tasks(self, due_filters=None, account_names=None):
        """Fetches Outlook Tasks from enabled accounts using safe Tables."""
        for attempt in range(2):
            if not self.namespace:
                 if not self.connect(): return []
            try:
                all_results = []
                
                for store in self._get_enabled_stores(account_names):
                    try:
                        tasks_folder = store.GetDefaultFolder(13)
                        
                        restricts = ["[Complete] = False"]
                        
                        # Date Filter Logic
                        if due_filters and len(due_filters) > 0:
                            date_queries = []
                            now = datetime.now()
                            today = now.replace(hour=0, minute=0, second=0, microsecond=0)
                            tomorrow = today + timedelta(days=1)
                            db_tomorrow = today + timedelta(days=2)
                            
                            for filter_name in due_filters:
                                if filter_name == "Overdue":
                                    date_queries.append("[DueDate] < '{}'".format(today.strftime('%d/%m/%Y %H:%M'))) 
                                elif filter_name == "Today":
                                    date_queries.append("([DueDate] >= '{}' AND [DueDate] < '{}')".format(today.strftime('%d/%m/%Y %H:%M'), tomorrow.strftime('%d/%m/%Y %H:%M')))
                                elif filter_name == "Tomorrow":
                                    date_queries.append("([DueDate] >= '{}' AND [DueDate] < '{}')".format(tomorrow.strftime('%d/%m/%Y %H:%M'), db_tomorrow.strftime('%d/%m/%Y %H:%M')))
                                elif filter_name == "Next 7 Days":
                                    next_week = today + timedelta(days=8)
                                    date_queries.append("([DueDate] >= '{}' AND [DueDate] < '{}')".format(today.strftime('%d/%m/%Y %H:%M'), next_week.strftime('%d/%m/%Y %H:%M')))
                                elif filter_name == "No Date":
                                    date_queries.append("([DueDate] IS NULL OR [DueDate] > '01/01/4500')")
                            
                            if date_queries:
                                combined_date_query = " OR ".join(date_queries)
                                restricts.append("({})".format(combined_date_query))

                        restrict_str = " AND ".join(restricts) if restricts else ""
                        
                        try:
                            table = tasks_folder.GetTable(restrict_str) if restrict_str else tasks_folder.GetTable()
                        except:
                            continue

                        table.Columns.RemoveAll()
                        table.Columns.Add("Subject")
                        table.Columns.Add("DueDate")
                        table.Columns.Add("EntryID")
                        
                        count = 0
                        while not table.EndOfTable and count < 30:
                            row = table.GetNextRow()
                            if not row: break
                            
                            try:
                                vals = row.GetValues()
                                
                                all_results.append({
                                    "subject": vals[0],
                                    "due": vals[1],
                                    "entry_id": vals[2],
                                    "is_task": True,
                                    "account": store.DisplayName,
                                    "store_id": store.StoreID
                                })
                                count += 1
                            except:
                                continue
                    except:
                        continue # Skip store if tasks failed
                        
                # Sort combined results
                all_results.sort(key=lambda x: x["due"].timestamp() if getattr(x["due"], 'timestamp', None) else 0)
                
                return all_results
            except Exception as e:
                print("Tasks error: {}".format(e))
                self.namespace = None
        return []

    def get_inbox_items(self, count=20, unread_only=False, only_flagged=False, due_filters=None, account_names=None, account_config=None):
        """Fetches items from configured folders for enabled accounts."""
        for attempt in range(2):
            if not self.namespace:
                if not self.connect(): return [], 0

            # Detect stale COM connection (e.g. after network change)
            if not self._is_connection_healthy():
                print("COM connection stale in get_inbox_items, reconnecting...")
                if not self.reconnect(): return [], 0

            try:
                all_items = []
                total_unread_count = 0
                
                for store in self._get_enabled_stores(account_names):
                    try:
                        # Determine folders to scan
                        folders_to_scan = []
                        
                        # Check config for this account
                        if account_config and store.DisplayName in account_config:
                            conf = account_config[store.DisplayName]
                            if "email_folders" in conf and conf["email_folders"]:
                                for path in conf["email_folders"]:
                                    f = self.get_folder_by_path(store, path)
                                    if f: folders_to_scan.append(f)
                        
                        # Fallback to Inbox if no specific folders configured
                        if not folders_to_scan:
                            try:
                                folders_to_scan.append(store.GetDefaultFolder(6))
                            except: pass
                            
                        for folder in folders_to_scan:
                             try:
                                 total_unread_count += folder.UnReadItemCount
                             except: pass
                             items = self._fetch_items_from_inbox_folder(folder, count, unread_only, only_flagged, due_filters, store)
                             all_items.extend(items)
                    except:
                        continue
                        
                # Sort merged results by ReceivedTime (Descending)
                def sort_key(x):
                    dt = x.get("received_dt")
                    if dt:
                        return dt
                    return datetime.min

                all_items.sort(key=sort_key, reverse=True)
                
                return all_items[:count], total_unread_count
                
            except Exception as e:
                self._log_debug("Inbox error: {}".format(e))
                print("Inbox error: {}".format(e))
                
        return [], 0

    def _log_debug(self, msg):
        """Log debug messages to AppData for troubleshooting frozen builds."""
        try:
            app_data = os.path.join(os.environ.get("LOCALAPPDATA", "."), "OutlookSidebar")
            if not os.path.exists(app_data):
                os.makedirs(app_data)
            with open(os.path.join(app_data, "debug_outlook.log"), "a") as f:
                f.write("{} - {}\n".format(datetime.now(), msg))
        except:
            pass

    def _fetch_items_from_inbox_folder(self, folder, count, unread_only, only_flagged, due_filters, store):
        """Helper to fetch items from a single inbox folder."""
        restricts = []
        
        if only_flagged:
            restricts.append("[FlagStatus] <> 0")
            
            if due_filters and len(due_filters) > 0:
                date_queries = []
                now = datetime.now()
                today = now.replace(hour=0, minute=0, second=0, microsecond=0)
                tomorrow = today + timedelta(days=1)
                db_tomorrow = today + timedelta(days=2)
                
                for filter_name in due_filters:
                    if filter_name == "Overdue":
                        date_queries.append("[TaskDueDate] < '{}'".format(today.strftime('%d/%m/%Y %H:%M'))) 
                    elif filter_name == "Today":
                        date_queries.append("([TaskDueDate] >= '{}' AND [TaskDueDate] < '{}')".format(today.strftime('%d/%m/%Y %H:%M'), tomorrow.strftime('%d/%m/%Y %H:%M')))
                    elif filter_name == "Tomorrow":
                        date_queries.append("([TaskDueDate] >= '{}' AND [TaskDueDate] < '{}')".format(tomorrow.strftime('%d/%m/%Y %H:%M'), db_tomorrow.strftime('%d/%m/%Y %H:%M')))
                
                if date_queries:
                    combined_date_query = " OR ".join(date_queries)
                    restricts.append("({})".format(combined_date_query))

        if unread_only:
            # Use DASL for safer unread filtering (avoids some Table API bugs with [UnRead])
            restricts.append("@SQL=\"urn:schemas:httpmail:read\" = 0")
        else:
            # Use DASL for safer date format (US format MM/DD/YYYY) to avoid locale issues
            # Limit scan to recent 7 days
            cutoff = (datetime.now() - timedelta(days=7)).strftime('%m/%d/%Y %H:%M')
            restricts.append("@SQL=\"urn:schemas:httpmail:datereceived\" >= '{}'".format(cutoff))
        
        restrict_str = " AND ".join(restricts) if restricts else ""
        
        try:
            # Log the restriction string for debugging
            # self._log_debug("Fetch from {}: Restrict='{}'".format(folder.Name, restrict_str))
            
            # Table approach for safety and speed
            table = folder.GetTable(restrict_str) if restrict_str else folder.GetTable()
            table.Sort("ReceivedTime", True)
            
            # Remove all columns and add only what we need
            # NOTE: "Body" is NOT supported by the Table API and causes
            # GetValues() to fail with "The parameter is incorrect".
            table.Columns.RemoveAll()
            table.Columns.Add("EntryID")
            table.Columns.Add("Subject")
            table.Columns.Add("SenderName")
            table.Columns.Add("ReceivedTime")
            table.Columns.Add("UnRead")
            table.Columns.Add("FlagStatus")
            try: table.Columns.Add("MessageClass")
            except: pass
            try: table.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x0E1B000B")  # PR_HASATTACH - real attachments only
            except: pass
            try: table.Columns.Add("Importance")
            except: pass
            
            items = []
            c = 0
            while not table.EndOfTable and c < count:
                try:
                    row = table.GetNextRow()
                    if not row: break
                    
                    vals = row.GetValues()
                    # EntryID=0, Subject=1, Sender=2, Recv=3, UnRead=4, Flag=5, Class=6, HasAttach=7, Importance=8
                    
                    # Filter out Non-Mail items if possible (e.g. Meeting Requests/Responses often clog inbox)
                    msg_class = vals[6] if len(vals) > 6 else "IPM.Note"
                    has_attach = vals[7] if len(vals) > 7 else False
                    importance = vals[8] if len(vals) > 8 else 1
                    
                    items.append({
                        "entry_id": vals[0],
                        "subject": vals[1],
                        "sender": vals[2],
                        "received_dt": vals[3],
                        "unread": vals[4],
                        "flag_status": vals[5],
                        "has_attachments": bool(has_attach),
                        "importance": importance,
                        "preview": "",
                        "is_meeting_request": "IPM.Schedule" in str(msg_class),
                        "store_id": store.StoreID, # Needed for actions
                        "account": store.DisplayName
                    })
                    c += 1
                except Exception as row_err:
                    self._log_debug("Fetch row error: {}".format(row_err))
                    continue
                
            return items
        except Exception as e:
            self._log_debug("Fetch error ({}): {}".format(folder.Name, e))
            return []

    def get_folder_list(self, account_name=None):
        """Returns a recursive list of folder paths for selection."""
        if not self.namespace: return []
        paths = []
        
        try:
            for store in self._get_enabled_stores([account_name] if account_name else None):
                root = store.GetRootFolder()
                self._recurse_folders(root, "", paths)
        except Exception as e:
            print("Error getting folders: {}".format(e))
            
        return paths
        
    def _recurse_folders(self, folder, current_path, paths):
        try:
            for f in folder.Folders:
                # Build path
                p = f.Name if not current_path else current_path + "/" + f.Name
                paths.append(p)
                # Recurse
                if f.Folders.Count > 0:
                    self._recurse_folders(f, p, paths)
        except: pass

    def get_item_by_entryid(self, entry_id, store_id=None):
        """Retrieves an Outlook item by its EntryID (and optional StoreID)."""
        if not self.namespace:
            if not self.connect(): return None
        try:
            if store_id:
                return self.namespace.GetItemFromID(entry_id, store_id)
            else:
                return self.namespace.GetItemFromID(entry_id)
        except Exception as e:
            print("Error getting item by EntryID: {}".format(e))
            return None

    def find_folder_by_name(self, folder_path):
        """Searches all stores for a folder matching the given path (e.g. 'Deleted Items' or 'Inbox/Subfolder')."""
        if not self.namespace:
            if not self.connect(): return None
        try:
            for store in self.namespace.Stores:
                result = self.get_folder_by_path(store, folder_path)
                if result:
                    return result
        except Exception as e:
            print("Error finding folder '{}': {}".format(folder_path, e))
        return None

    def find_folder_in_store(self, store_id, folder_path):
        """Finds a folder by path within a specific store (by StoreID)."""
        if not self.namespace:
            if not self.connect(): return None
        try:
            for store in self.namespace.Stores:
                if store.StoreID == store_id:
                    return self.get_folder_by_path(store, folder_path)
        except Exception as e:
            print("Error finding folder '{}' in store: {}".format(folder_path, e))
        return None

    def mark_task_complete(self, entry_id, store_id=None):
        """Marks an Outlook Task as complete."""
        try:
            if store_id:
                item = self.namespace.GetItemFromID(entry_id, store_id)
            else:
                item = self.namespace.GetItemFromID(entry_id)
            # TaskItem.MarkComplete() sets Status=2 (Complete) and PercentComplete=100
            item.MarkComplete()
            return True
        except Exception as e:
            err_msg = str(e)
            # "already complete" is still a success from the user's perspective
            if "already complete" in err_msg.lower():
                return True
            print("Error marking task complete: {}".format(e))
            return False

    # --- Actions ---
    
    def mark_as_read(self, entry_id, store_id=None):
        try:
            if store_id:
                item = self.namespace.GetItemFromID(entry_id, store_id)
            else:
                item = self.namespace.GetItemFromID(entry_id)
            item.UnRead = False
            item.Save()
            return True
        except: return False

    def toggle_flag(self, entry_id, store_id=None):
        try:
            if store_id:
                item = self.namespace.GetItemFromID(entry_id, store_id)
            else:
                item = self.namespace.GetItemFromID(entry_id)
            
            # Simple toggle: If flagged, unflag. If not, flag for today.
            if item.FlagStatus != 0:
                item.FlagStatus = 0 # olNoFlag
            else:
                item.FlagStatus = 2 # olFlagMarked
                # Default to today?
                # item.TaskStartDate = datetime.now()
            item.Save()
            return True
        except: return False

    def unflag_email(self, entry_id, store_id=None):
        try:
            if store_id:
                item = self.namespace.GetItemFromID(entry_id, store_id)
            else:
                item = self.namespace.GetItemFromID(entry_id)
            # Use ClearTaskFlag to properly remove the flag (matches email card behavior)
            if item.IsMarkedAsTask:
                item.ClearTaskFlag()
            item.FlagStatus = 0  # olNoFlag - belt and braces
            item.Save()
            return True
        except Exception as e:
            print("Unflag error: {}".format(e))
            return False

    def delete_email(self, entry_id, store_id=None):
        try:
            if store_id:
                item = self.namespace.GetItemFromID(entry_id, store_id)
            else:
                item = self.namespace.GetItemFromID(entry_id)
            item.Delete()
            return True
        except: return False
    
    def open_item(self, entry_id, store_id=None):
        try:
            if store_id:
                item = self.namespace.GetItemFromID(entry_id, store_id)
            else:
                item = self.namespace.GetItemFromID(entry_id)
            item.Display()
            # Bring to front?
            return True
        except Exception as e:
            print("Open error: {}".format(e))
            return False

    def send_email_with_attachment(self, recipient, subject, body, attachment_path):
        """Send an email with a file attachment via Outlook."""
        try:
            mail = self.outlook.CreateItem(0)  # olMailItem
            mail.To = recipient
            mail.Subject = subject
            mail.Body = body
            mail.Attachments.Add(attachment_path)
            mail.Send()
            return True
        except Exception as e:
            print("Error sending email: {}".format(e))
            return False

    def create_email(self):
        try:
            mail = self.outlook.CreateItem(0) # olMailItem
            mail.Display()
        except: pass

    def create_appointment(self):
        try:
            appt = self.outlook.CreateItem(1) # olAppointmentItem
            appt.Display()
        except: pass

    def create_meeting(self):
        try:
            meeting = self.outlook.CreateItem(1)  # olAppointmentItem
            meeting.MeetingStatus = 1  # olMeeting - enables attendee picker
            meeting.Display()
        except: pass
        
    def create_task(self):
        try:
            task = self.outlook.CreateItem(3) # olTaskItem
            task.Display()
        except: pass
        
    def create_contact(self):
        try:
             contact = self.outlook.CreateItem(2) # olContactItem
             contact.Display()
        except: pass

    def complete_task(self, entry_id):
        try:
            item = self.namespace.GetItemFromID(entry_id)
            item.MarkComplete()
            return True
        except: return False

    def dismiss_reminder(self, entry_id):
        """Dismissing is harder via OOM for reminders specifically if not firing.
           But for calendar items/tasks we can just remove/complete check."""
        # For simplicity, let's just Open it so user can dismiss? 
        # API for Dismissing active reminders is Application.Reminders...
        try:
             # Look in reminders collection
             for rem in self.outlook.Reminders:
                 try:
                     if rem.Item.EntryID == entry_id:
                         rem.Dismiss()
                         return True
                 except: continue
        except: pass
        return False

    def search_contacts(self, query, max_results=8):
        """Search Outlook Contacts and GAL for matching names/emails.
        
        Returns list of dicts: [{"name": "...", "email": "..."}, ...]
        """
        if not query or len(query) < 2:
            return []
        
        if not self.namespace:
            if not self.connect():
                return []
        
        results = []
        seen_emails = set()
        query_lower = query.lower()
        
        # 1. Search default Contacts folder (fast, Table API)
        try:
            for store in self.namespace.Stores:
                try:
                    contacts_folder = store.GetDefaultFolder(10)  # olFolderContacts
                    table = contacts_folder.GetTable()
                    table.Columns.RemoveAll()
                    table.Columns.Add("FullName")
                    table.Columns.Add("Email1Address")
                    
                    while not table.EndOfTable and len(results) < max_results:
                        row = table.GetNextRow()
                        if not row: break
                        vals = row.GetValues()
                        name = vals[0] or ""
                        email = vals[1] or ""
                        
                        if not email:
                            continue
                        
                        # Match on name or email
                        if query_lower in name.lower() or query_lower in email.lower():
                            email_lower = email.lower()
                            if email_lower not in seen_emails:
                                seen_emails.add(email_lower)
                                results.append({"name": name, "email": email})
                except:
                    continue
        except:
            pass
        
        # 2. Search GAL (Global Address List) if available
        if len(results) < max_results:
            try:
                for addr_list in self.namespace.AddressLists:
                    if addr_list.AddressListType == 1:  # olExchangeGlobalAddressList
                        entries = addr_list.AddressEntries
                        count = 0
                        for entry in entries:
                            if len(results) >= max_results:
                                break
                            count += 1
                            if count > 500:  # Safety limit for large GALs
                                break
                            try:
                                name = entry.Name or ""
                                if query_lower not in name.lower():
                                    continue
                                # Get SMTP address
                                email = ""
                                try:
                                    eu = entry.GetExchangeUser()
                                    if eu:
                                        email = eu.PrimarySmtpAddress or ""
                                except:
                                    pass
                                
                                if email:
                                    email_lower = email.lower()
                                    if email_lower not in seen_emails:
                                        seen_emails.add(email_lower)
                                        results.append({"name": name, "email": email})
                            except:
                                continue
                        break  # Only search the first GAL
            except:
                pass
        
        return results[:max_results]

    def get_category_map(self):
        """Returns a dict {CategoryName: ColorIndex}."""
        cat_map = {}
        if not self.namespace: return {}
        try:
             # categories might not be available if not connected properly?
             # Categories is a collection on NameSpace
             for cat in self.namespace.Categories:
                 cat_map[cat.Name] = cat.Color
        except: pass
        return cat_map

    def get_pulse_status(self, account_names=None):
        """Returns a lightweight status dict for the pulse indicator.
        
        Returns: {"calendar": "Today"|None, "tasks": "Overdue"|"Today"|None}
        """
        result = {"calendar": None, "tasks": None}
        
        if not self.namespace:
            if not self.connect():
                return result
        
        now = datetime.now()
        today = now.replace(hour=0, minute=0, second=0, microsecond=0)
        tomorrow = today + timedelta(days=1)
        
        # Check calendar — any items today?
        try:
            s_str = today.strftime('%d/%m/%Y %H:%M')
            e_str = tomorrow.strftime('%d/%m/%Y %H:%M')
            
            for store in self._get_enabled_stores(account_names):
                try:
                    cal = store.GetDefaultFolder(9)
                    items = cal.Items
                    items.Sort("[Start]")
                    items.IncludeRecurrences = True
                    restrict = "[Start] >= '{}' AND [Start] <= '{}'".format(s_str, e_str)
                    filtered = items.Restrict(restrict)
                    # Just check if any exist
                    if filtered.Count > 0:
                        result["calendar"] = "Today"
                        break
                except:
                    continue
        except:
            pass
        
        # Check tasks — any overdue or due today?
        try:
            for store in self._get_enabled_stores(account_names):
                try:
                    tasks_folder = store.GetDefaultFolder(13)
                    
                    # Check overdue first (higher priority)
                    overdue_q = "[Complete] = False AND [DueDate] < '{}'".format(
                        today.strftime('%d/%m/%Y %H:%M'))
                    try:
                        table = tasks_folder.GetTable(overdue_q)
                        if not table.EndOfTable:
                            result["tasks"] = "Overdue"
                            break
                    except:
                        pass
                    
                    # Check today
                    today_q = "[Complete] = False AND [DueDate] >= '{}' AND [DueDate] < '{}'".format(
                        today.strftime('%d/%m/%Y %H:%M'),
                        tomorrow.strftime('%d/%m/%Y %H:%M'))
                    try:
                        table = tasks_folder.GetTable(today_q)
                        if not table.EndOfTable:
                            result["tasks"] = "Today"
                            break
                    except:
                        pass
                except:
                    continue
        except:
            pass
        
        return result

    def get_due_status(self, due_date):
        """Helper to return Today/Tomorrow/Overdue status string."""
        if not due_date: return None
        try:
            now = datetime.now()
            today = now.replace(hour=0, minute=0, second=0, microsecond=0)
            tomorrow = today + timedelta(days=1)
            
            # Normalize due_date
            dd = due_date
            if getattr(dd, "tzinfo", None):
                 dd = dd.replace(tzinfo=None)
            
            if dd < today:
                return "Overdue"
            elif dd >= today and dd < tomorrow:
                return "Today"
            elif dd >= tomorrow and dd < (tomorrow + timedelta(days=1)):
                return "Tomorrow"
            else:
                return "Later"
        except:
            return None
