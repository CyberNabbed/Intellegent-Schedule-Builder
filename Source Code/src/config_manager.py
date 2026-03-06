import json
import os

CONFIG_FILE = "team_config.json"

# Default settings for first-time setup
DEFAULT_SETTINGS = {
    "email_method": None,         # None | "outlook_com"
    "email_method_verified": False # Has the user confirmed Outlook COM works?
}


class ConfigManager:
    def __init__(self):
        # Use LOCALAPPDATA on Windows, fallback to ~/.config on Mac/Linux
        app_data_dir = os.environ.get('LOCALAPPDATA')
        if not app_data_dir:
            app_data_dir = os.path.expanduser('~/.config')
            
        self.config_dir = os.path.join(app_data_dir, 'ScheduleBuilder')
        os.makedirs(self.config_dir, exist_ok=True)
        self.config_path = os.path.join(self.config_dir, CONFIG_FILE)
        self.team_data = self.load_config()

    def load_config(self):
        """Loads the team configuration from JSON."""
        if not os.path.exists(self.config_path):
            return None  # Return None to signal "First Run"
        try:
            with open(self.config_path, 'r') as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError, PermissionError) as e:
            print(f"Warning: Failed to load config file: {e}")
            return None

    def save_config(self, employees):
        """
        Saves the list of employees to JSON, preserving existing settings.
        Format: [{"name": "Name", "type": "early" | "late", "email": "..."}]
        """
        # Preserve existing settings when saving employees
        existing_settings = self.get_settings()
        data = {
            "employees": employees,
            "settings": existing_settings
        }
        self._write_config(data)

    def _write_config(self, data):
        """Low-level write helper — writes the full config dict to disk."""
        try:
            config_dir = os.path.dirname(self.config_path)
            if not os.access(config_dir, os.W_OK):
                raise PermissionError(f"No write permission for directory: {config_dir}")

            with open(self.config_path, 'w') as f:
                json.dump(data, f, indent=4)
            self.team_data = data
        except (PermissionError, IOError) as e:
            raise Exception(f"Failed to save configuration: {e}")

    def get_employees(self):
        """Returns list of employee dicts. Safely handles missing email field."""
        if self.team_data:
            employees = self.team_data.get("employees", [])
            # Ensure every employee dict has an 'email' key (backward compat)
            for emp in employees:
                emp.setdefault("email", "")
            return employees
        return []

    def get_settings(self):
        """Returns the settings dict, falling back to defaults."""
        if self.team_data:
            saved = self.team_data.get("settings", {})
            # Merge with defaults so new keys are always present
            merged = dict(DEFAULT_SETTINGS)
            merged.update(saved)
            return merged
        return dict(DEFAULT_SETTINGS)

    def save_settings(self, settings):
        """
        Persists settings without touching the employee list.
        """
        employees = self.get_employees()
        data = {
            "employees": employees,
            "settings": settings
        }
        self._write_config(data)

    # ── Session persistence ──────────────────────────────────────────────

    @property
    def session_path(self):
        return os.path.join(self.config_dir, "saved_session.json")

    def save_session(self, data):
        """Save a session snapshot to disk. data must be JSON-serializable."""
        try:
            with open(self.session_path, 'w') as f:
                json.dump(data, f, indent=2)
        except (PermissionError, IOError) as e:
            raise Exception(f"Failed to save session: {e}")

    def load_session(self):
        """Load a previously saved session, or return None."""
        if not os.path.exists(self.session_path):
            return None
        try:
            with open(self.session_path, 'r') as f:
                return json.load(f)
        except (json.JSONDecodeError, PermissionError, IOError):
            return None

    def delete_session(self):
        """Remove the saved session file."""
        try:
            if os.path.exists(self.session_path):
                os.remove(self.session_path)
        except OSError:
            pass
