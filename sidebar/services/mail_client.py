import abc

class MailClient(abc.ABC):
    """Abstract interface that both COM and Graph backends implement."""
    
    # --- Connection ---
    @abc.abstractmethod
    def connect(self) -> bool:
        """Connects or authenticates to the mail backend."""
        pass
    
    @abc.abstractmethod
    def reconnect(self) -> bool:
        """Forces a reconnection / refresh."""
        pass
    
    @abc.abstractmethod
    def is_connected(self) -> bool:
        """Checks if currently connected and healthy."""
        pass
    
    # --- Accounts ---
    @abc.abstractmethod
    def get_accounts(self) -> list:
        """Returns a list of available account names/emails."""
        pass
    
    # --- Email ---
    @abc.abstractmethod
    def get_inbox_items(self, count=20, unread_only=False, only_flagged=False, 
                        due_filters=None, account_names=None, account_config=None) -> tuple:
        """
        Returns (emails, unread_count).
        Each email is a dict: { entry_id, subject, sender, received, unread, flag_status, ... }
        """
        pass
    
    @abc.abstractmethod
    def get_unread_count(self, account_names=None, account_config=None) -> int:
        """Returns the total number of unread emails."""
        pass
    
    @abc.abstractmethod
    def mark_as_read(self, entry_id, store_id=None) -> bool:
        """Marks an email as read."""
        pass
    
    @abc.abstractmethod
    def delete_email(self, entry_id, store_id=None) -> bool:
        """Moves an email to the Trash/Deleted Items folder."""
        pass
    
    @abc.abstractmethod
    def toggle_flag(self, entry_id, store_id=None) -> bool:
        """Toggles the flag status of an email."""
        pass
    
    @abc.abstractmethod
    def unflag_email(self, entry_id, store_id=None) -> bool:
        """Removes the flag from an email."""
        pass
    
    @abc.abstractmethod
    def open_item(self, entry_id, store_id=None):
        """Opens the item (e.g. in Outlook or web browser)."""
        pass
    
    # --- Calendar ---
    @abc.abstractmethod
    def get_calendar_items(self, start_dt, end_dt, account_names=None) -> list:
        """
        Returns a list of meetings between start_dt and end_dt.
        Each meeting is a dict: { entry_id, subject, start, end, location, response_status, ... }
        """
        pass
    
    # --- Tasks ---
    @abc.abstractmethod
    def get_tasks(self, due_filters=None, account_names=None) -> list:
        """
        Returns a list of tasks matching due_filters.
        Each task is a dict: { entry_id, subject, due, categories, ... }
        """
        pass
    
    @abc.abstractmethod
    def mark_task_complete(self, entry_id, store_id=None) -> bool:
        """Marks a task as complete."""
        pass
    
    # --- Quick Create ---
    @abc.abstractmethod
    def create_email(self):
        """Opens a 'New Email' compose window."""
        pass
    
    @abc.abstractmethod
    def create_meeting(self):
        """Opens a 'New Meeting' window."""
        pass
    
    @abc.abstractmethod
    def create_task(self):
        """Opens a 'New Task' window."""
        pass
    
    @abc.abstractmethod
    def create_contact(self):
         """Opens a 'New Contact' window."""
         pass
         
    # --- New Mail Detection ---
    @abc.abstractmethod
    def check_new_mail(self, account_names=None) -> bool:
        """Returns True if there is new mail since the last check."""
        pass
    
    # --- Pulse ---
    @abc.abstractmethod
    def get_pulse_status(self, account_names=None) -> dict:
        """
        Returns a status dict for the pulse notification.
        e.g., {"calendar": "Today", "tasks": "Overdue"}
        """
        pass
    
    # --- Utility ---
    @abc.abstractmethod
    def get_category_map(self) -> dict:
        """Returns a dict mapping CategoryName to ColorIndex/Hex."""
        pass
    
    @abc.abstractmethod
    def search_contacts(self, query, max_results=8) -> list:
        """
        Searches contacts/GAL for the query.
        Returns list of dicts: {"name": ..., "email": ...}
        """
        pass
    
    @abc.abstractmethod
    def get_folder_list(self, account_name=None) -> list:
        """Returns a hierarchical list of folder paths."""
        pass
        
    @abc.abstractmethod
    def get_native_app(self):
         """Returns native app ref if needed (e.g. COM 'outlook.Application')"""
         pass

    @abc.abstractmethod
    def send_email_with_attachment(self, recipient, subject, body, attachment_path) -> bool:
         """Sends a background email (used for Share feature)."""
         pass
