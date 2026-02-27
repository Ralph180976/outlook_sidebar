import os
import threading
import logging

try:
    import msal
except ImportError:
    msal = None

logger = logging.getLogger(__name__)

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
        self.app = None  # Set early so attribute always exists
        self.cache = None
        self._cache_path = None
        
        if msal is None:
            logger.warning("[GraphAuth] msal package not installed — Graph features disabled")
            return
        
        try:
            # Store token cache in Local AppData next to config
            cache_dir = os.path.join(os.environ.get("LOCALAPPDATA", "."), "InboxBar")
            os.makedirs(cache_dir, exist_ok=True)
            self._cache_path = os.path.join(cache_dir, "graph_token_cache.bin")
            
            self.cache = msal.SerializableTokenCache()
            if os.path.exists(self._cache_path):
                try:
                    with open(self._cache_path, "r") as f:
                        content = f.read()
                        if content:
                            self.cache.deserialize(content)
                except Exception as e:
                    print(f"[GraphAuth] Ignored corrupt cache: {e}")
            
            # Use 'common' for multi-tenant so any Microsoft account works
            # Configure a requests session with a short timeout to prevent
            # acquire_token_silent() from blocking for 30+ seconds when offline
            try:
                import requests
                http_session = requests.Session()
                # Monkey-patch the send method to enforce a connection timeout
                _original_send = http_session.send
                def _send_with_timeout(prepared, **kwargs):
                    kwargs.setdefault('timeout', (10, 30))
                    return _original_send(prepared, **kwargs)
                http_session.send = _send_with_timeout
            except ImportError:
                http_session = None
            
            self.app = msal.PublicClientApplication(
                CLIENT_ID,
                authority="https://login.microsoftonline.com/common",
                token_cache=self.cache,
                http_client=http_session,
            )
        except Exception as e:
            logger.error(f"[GraphAuth] Failed to initialize MSAL: {e}")
            self.app = None

    def get_token(self, interactive=False):
        """
        Attempts to get a valid access token.
        If interactive=True and silent auth fails, it will pop up a browser window.
        """
        if not self.app:
            return None
        
        accounts = self.app.get_accounts()
        
        if accounts:
            # Try silent token acquisition (refreshing if needed)
            # We just use the first account for now (simplest approach)
            try:
                result = self.app.acquire_token_silent(SCOPES, account=accounts[0])
                if result and "access_token" in result:
                    self._save_cache()
                    return result["access_token"]
            except Exception as e:
                # Network errors (DNS, connection timeout) during token refresh
                # Return None so the caller treats it as "not authenticated" rather than crashing
                print(f"[GraphAuth] Token refresh failed (likely offline): {e}")
                return None
        
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
        if not self.app:
            return []
        return self.app.get_accounts()

    def logout(self):
        """Signs the user out by clearing the token cache."""
        if not self.app:
            return
        for account in self.app.get_accounts():
            self.app.remove_account(account)
        self._save_cache()
        
    def get_current_user_email(self):
        """Returns the email address of the first logged in account, or None."""
        if not self.app:
            return None
        accounts = self.app.get_accounts()
        if accounts:
            return accounts[0].get("username")
        return None

    def _save_cache(self):
        """Saves the token cache to disk if it was modified."""
        if not self.cache or not self._cache_path:
            return
        if self.cache.has_state_changed:
            with open(self._cache_path, "w") as f:
                f.write(self.cache.serialize())
