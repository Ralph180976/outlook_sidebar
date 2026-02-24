import msal
import os
import threading

# Use the Client ID provided by the user
CLIENT_ID = "fd659fea-ed23-426f-9c82-1f61f6b40fb2"
SCOPES = [
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.ReadWrite",
    "Tasks.ReadWrite",
    "User.Read"
]

class GraphAuth:
    """Manages Microsoft Graph API authentication via MSAL."""
    
    _instance = None
    _lock = threading.Lock()

    def __new__(cls):
        with cls._lock:
            if cls._instance is None:
                cls._instance = super(GraphAuth, cls).__new__(cls)
                cls._instance._init()
            return cls._instance

    def _init(self):
        # Store token cache in Local AppData next to config
        cache_dir = os.path.join(os.environ.get("LOCALAPPDATA", "."), "InboxBar")
        os.makedirs(cache_dir, exist_ok=True)
        self._cache_path = os.path.join(cache_dir, "graph_token_cache.bin")
        
        self.cache = msal.SerializableTokenCache()
        if os.path.exists(self._cache_path):
            with open(self._cache_path, "r") as f:
                self.cache.deserialize(f.read())
        
        # Use 'common' for multi-tenant so any Microsoft account works
        self.app = msal.PublicClientApplication(
            CLIENT_ID,
            authority="https://login.microsoftonline.com/common",
            token_cache=self.cache,
        )

    def get_token(self, interactive=False):
        """
        Attempts to get a valid access token.
        If interactive=True and silent auth fails, it will pop up a browser window.
        """
        accounts = self.app.get_accounts()
        
        if accounts:
            # Try silent token acquisition (refreshing if needed)
            # We just use the first account for now (simplest approach)
            result = self.app.acquire_token_silent(SCOPES, account=accounts[0])
            if result and "access_token" in result:
                self._save_cache()
                return result["access_token"]
        
        # If silent fails and interactive is not requested, return None
        if not interactive:
            return None
            
        # Interactive login required (opens browser)
        print("[GraphAuth] Starting interactive login...")
        result = self.app.acquire_token_interactive(
            SCOPES,
            port=8400,
            prompt="select_account",
        )
        
        if "access_token" in result:
            self._save_cache()
            return result["access_token"]
            
        error_msg = result.get("error_description", result.get("error", "Unknown login error"))
        raise Exception(f"MSAL Login failed: {error_msg}")

    def get_accounts(self):
        """Returns the cached MSAL accounts (who is logged in)."""
        return self.app.get_accounts()

    def logout(self):
        """Signs the user out by clearing the token cache."""
        for account in self.app.get_accounts():
            self.app.remove_account(account)
        self._save_cache()
        
    def get_current_user_email(self):
        """Returns the email address of the first logged in account, or None."""
        accounts = self.app.get_accounts()
        if accounts:
            return accounts[0].get("username")
        return None

    def _save_cache(self):
        """Saves the token cache to disk if it was modified."""
        if self.cache.has_state_changed:
            with open(self._cache_path, "w") as f:
                f.write(self.cache.serialize())
