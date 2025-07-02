import requests
import json
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import configparser
from pathlib import Path
import logging
import subprocess
import socket
from typing import Dict, Any, Optional, Tuple, List


class Config:
    """Handles configuration management"""
    
    def __init__(self):
        self.config_folder = Path(os.getenv("APPDATA", "")) / "CCM"
        self.config_file = self.config_folder / "uploadsettings.ini"
        self.config = configparser.ConfigParser()
        self._ensure_config_exists()
        self._load_config()
    
    def _ensure_config_exists(self):
        """Create default config if it doesn't exist"""
        if not self.config_file.exists():
            self.config_folder.mkdir(parents=True, exist_ok=True)
            self.config['Paths'] = {'SaveFolderPath': ''}
            self.config['Companion'] = {'CompanionIP': ''}
            self.config['UI'] = {'ShowRefreshButton': 'False'}
            self._save_config()
            logging.info("Default config created")
    
    def _load_config(self):
        """Load configuration from file"""
        self.config.read(self.config_file)
    
    def _save_config(self):
        """Save configuration to file"""
        with open(self.config_file, 'w') as file:
            self.config.write(file)
    
    def get_save_folder(self) -> str:
        return self.config.get('Paths', 'SaveFolderPath', fallback='')
    
    def get_companion_ip(self) -> str:
        return self.config.get('Companion', 'CompanionIP', fallback='')
    
    def get_show_refresh_button(self) -> bool:
        return self.config.getboolean('UI', 'ShowRefreshButton', fallback=False)
    
    def update_settings(self, save_folder: str, companion_ip: str, show_refresh: bool = False):
        """Update configuration settings"""
        # Ensure all sections exist
        if not self.config.has_section('Paths'):
            self.config.add_section('Paths')
        if not self.config.has_section('Companion'):
            self.config.add_section('Companion')
        if not self.config.has_section('UI'):
            self.config.add_section('UI')
        
        self.config.set('Paths', 'SaveFolderPath', save_folder)
        self.config.set('Companion', 'CompanionIP', companion_ip)
        self.config.set('UI', 'ShowRefreshButton', str(show_refresh))
        self._save_config()
    
    def is_valid(self) -> bool:
        """Check if required configuration is present"""
        return bool(self.get_save_folder() and self.get_companion_ip())


class CompanionAPI:
    """Handles communication with Companion API"""
    
    def __init__(self, companion_ip: str):
        self.companion_ip = companion_ip
        self.base_url = f"http://{companion_ip}:8000/api/custom-variable"
        self.timeout = 5
        
        # Variable mappings
        self.variable_map = {
            'song1': '9amSong1',
            'song2': '9amSong2', 
            'song3': '9amSong3',
            'start': '9amStart',
            'end': '9amEnd',
            'communion': '9amCommunion',
            'song1path': '9amSong1Path',
            'song2path': '9amSong2Path',
            'song3path': '9amSong3Path',
            'startpath': '9amStartPath',
            'endpath': '9amEndPath',
            'communionpath': '9amCommunionPath'
        }
    
    def test_connection(self) -> bool:
        """Test if Companion is reachable at the configured IP"""
        try:
            # First check if the port is open
            with socket.create_connection((self.companion_ip, 8000), timeout=3):
                pass
            
            # Then try to make an API call
            url = f"{self.base_url}/ServiceDate/value"
            response = requests.get(url, timeout=3)
            return response.status_code == 200
        except (socket.error, requests.RequestException):
            return False
    
    def get_current_service_date(self) -> Optional[str]:
        """Fetch current service date from Companion"""
        try:
            url = f"{self.base_url}/ServiceDate/value"
            response = requests.get(url, timeout=self.timeout)
            response.raise_for_status()
            return response.text.strip()
        except requests.RequestException as e:
            logging.error(f"Failed to fetch current service date: {e}")
            return None
    
    def update_service_data(self, song_data: Dict[str, str], service_date: str) -> List[str]:
        """Send service data to Companion API"""
        errors = []
        
        # Update service date first
        try:
            url = f"{self.base_url}/ServiceDate/value?value={service_date}"
            requests.post(url, timeout=self.timeout)
        except requests.RequestException as e:
            errors.append(f"ServiceDate: {e}")
        
        # Update song data
        for key, value in song_data.items():
            var_name = self.variable_map.get(key)
            if var_name:
                try:
                    url = f"{self.base_url}/{var_name}/value?value={value}"
                    requests.post(url, timeout=self.timeout)
                except requests.RequestException as e:
                    errors.append(f"{var_name}: {e}")
        
        return errors


class VLCLauncher:
    """Handles VLC launching functionality"""
    
    @staticmethod
    def find_vlc_path() -> Optional[str]:
        """Find VLC installation path on Windows"""
        common_paths = [
            r"C:\Program Files\VideoLAN\VLC\vlc.exe",
            r"C:\Program Files (x86)\VideoLAN\VLC\vlc.exe",
            r"C:\Users\{}\AppData\Local\Programs\VideoLAN\VLC\vlc.exe".format(os.getenv('USERNAME', '')),
        ]
        
        for path in common_paths:
            if os.path.exists(path):
                return path
        
        # Try to find in PATH
        try:
            subprocess.run(["vlc", "--version"], capture_output=True, timeout=3)
            return "vlc"
        except (subprocess.TimeoutExpired, subprocess.CalledProcessError, FileNotFoundError):
            pass
        
        return None
    
    @staticmethod
    def launch_vlc() -> bool:
        """Launch VLC media player"""
        vlc_path = VLCLauncher.find_vlc_path()
        
        if not vlc_path:
            return False
        
        try:
            subprocess.Popen([vlc_path], shell=False)
            return True
        except Exception as e:
            logging.error(f"Failed to launch VLC: {e}")
            return False


class ServiceDataManager:
    """Manages service data operations"""
    
    def __init__(self, json_path: str):
        self.json_path = Path(json_path)
        self.data = {}
        self.load_data()
    
    def load_data(self) -> bool:
        """Load service data from JSON file"""
        try:
            if self.json_path.exists():
                with open(self.json_path, 'r', encoding='utf-8') as file:
                    self.data = json.load(file)
                logging.info("Service data loaded successfully")
                return True
            else:
                logging.warning(f"JSON file not found: {self.json_path}")
                return False
        except Exception as e:
            logging.error(f"Failed to load service data: {e}")
            return False
    
    def get_sorted_dates(self) -> List[str]:
        """Get sorted list of service dates"""
        return sorted(self.data.keys(), key=lambda x: datetime.strptime(x, '%d/%m/%Y'))
    
    def find_nearest_upcoming_service(self) -> Optional[Tuple[str, Dict[str, Any]]]:
        """Find the nearest upcoming service"""
        today = datetime.today().date()
        future_events = [
            (date_str, details) 
            for date_str, details in self.data.items() 
            if datetime.strptime(date_str, '%d/%m/%Y').date() >= today
        ]
        
        if future_events:
            future_events.sort(key=lambda x: datetime.strptime(x[0], '%d/%m/%Y').date())
            return future_events[0]
        return None
    
    def extract_song_data(self, event_data: Dict[str, Any]) -> Dict[str, str]:
        """Extract song data from event data"""
        song_keys = [
            'song1', 'song2', 'song3', 'start', 'end', 'communion',
            'song1path', 'song2path', 'song3path', 'startpath', 'endpath', 'communionpath'
        ]
        return {key: event_data.get(key, 'none') for key in song_keys}
    
    def get_service_data(self, date_str: str) -> Dict[str, str]:
        """Get service data for a specific date"""
        event_data = self.data.get(date_str, {})
        return self.extract_song_data(event_data)


class CompanionConnectionDialog:
    """Dialog for updating Companion IP when connection fails"""
    
    def __init__(self, parent, current_ip: str, config: Config):
        self.config = config
        self.new_ip = None
        self.window = tk.Toplevel(parent)
        self.setup_ui(current_ip)
    
    def setup_ui(self, current_ip: str):
        """Setup the connection dialog UI"""
        self.window.title("Companion Connection Failed")
        self.window.geometry("400x200")
        self.window.resizable(False, False)
        self.window.grab_set()  # Make dialog modal
        
        # Center the dialog
        self.window.transient(self.window.master)
        
        # Warning message
        ttk.Label(
            self.window, 
            text="⚠️ Cannot connect to Companion",
            font=("Arial", 12, "bold"),
            foreground="red"
        ).pack(pady=10)
        
        ttk.Label(
            self.window,
            text=f"Current IP: {current_ip}",
            font=("Arial", 10)
        ).pack(pady=5)
        
        ttk.Label(
            self.window,
            text="Please enter the correct Companion IP address:",
            font=("Arial", 10)
        ).pack(pady=5)
        
        # IP input
        self.ip_var = tk.StringVar(value=current_ip)
        ip_frame = ttk.Frame(self.window)
        ip_frame.pack(pady=10)
        
        ttk.Label(ip_frame, text="IP Address:").pack(side="left", padx=5)
        self.ip_entry = ttk.Entry(ip_frame, textvariable=self.ip_var, width=20)
        self.ip_entry.pack(side="left", padx=5)
        self.ip_entry.focus()
        self.ip_entry.select_range(0, tk.END)
        
        # Buttons
        button_frame = ttk.Frame(self.window)
        button_frame.pack(pady=20)
        
        ttk.Button(
            button_frame, 
            text="Test & Save", 
            command=self.test_and_save
        ).pack(side="left", padx=5)
        
        ttk.Button(
            button_frame, 
            text="Skip", 
            command=self.skip
        ).pack(side="left", padx=5)
        
        # Bind Enter key to test and save
        self.window.bind('<Return>', lambda e: self.test_and_save())
    
    def test_and_save(self):
        """Test the new IP and save if successful"""
        new_ip = self.ip_var.get().strip()
        
        if not new_ip:
            messagebox.showerror("Error", "Please enter an IP address")
            return
        
        # Test connection
        try:
            test_api = CompanionAPI(new_ip)
            if test_api.test_connection():
                # Save the new IP
                self.config.update_settings(
                    self.config.get_save_folder(),
                    new_ip,
                    self.config.get_show_refresh_button()
                )
                self.new_ip = new_ip
                messagebox.showinfo("Success", "Connection successful! IP address saved.")
                self.window.destroy()
            else:
                messagebox.showerror(
                    "Connection Failed", 
                    f"Still cannot connect to Companion at {new_ip}.\n"
                    "Please check:\n"
                    "• Companion is running\n"
                    "• IP address is correct\n"
                    "• Port 8000 is accessible"
                )
        except Exception as e:
            messagebox.showerror("Error", f"Connection test failed: {e}")
    
    def skip(self):
        """Skip the connection test and continue"""
        self.window.destroy()
    
    def get_result(self) -> Optional[str]:
        """Get the new IP if successful"""
        self.window.wait_window()
        return self.new_ip


class SettingsDialog:
    """Settings dialog window"""
    
    def __init__(self, parent, config: Config, refresh_callback):
        self.config = config
        self.refresh_callback = refresh_callback
        self.window = tk.Toplevel(parent)
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the settings UI"""
        self.window.title("Settings")
        self.window.geometry("500x200")
        self.window.resizable(False, False)
        
        # Save folder path
        ttk.Label(self.window, text="Save Folder Path:").grid(
            row=0, column=0, padx=10, pady=10, sticky="w"
        )
        self.save_folder_var = tk.StringVar(value=self.config.get_save_folder())
        save_folder_frame = ttk.Frame(self.window)
        save_folder_frame.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        
        self.save_folder_entry = ttk.Entry(save_folder_frame, textvariable=self.save_folder_var, width=40)
        self.save_folder_entry.pack(side="left", fill="x", expand=True)
        
        ttk.Button(save_folder_frame, text="Browse", command=self.browse_folder).pack(side="right", padx=(5, 0))
        
        # Companion IP
        ttk.Label(self.window, text="Companion IP:").grid(
            row=1, column=0, padx=10, pady=10, sticky="w"
        )
        self.companion_ip_var = tk.StringVar(value=self.config.get_companion_ip())
        ttk.Entry(self.window, textvariable=self.companion_ip_var, width=40).grid(
            row=1, column=1, padx=10, pady=10, sticky="ew"
        )
        
        # Show refresh button checkbox
        self.show_refresh_var = tk.BooleanVar(value=self.config.get_show_refresh_button())
        ttk.Checkbutton(
            self.window, 
            text="Show Refresh Data Button", 
            variable=self.show_refresh_var
        ).grid(row=2, columnspan=2, pady=10)
        
        # Buttons
        button_frame = ttk.Frame(self.window)
        button_frame.grid(row=3, columnspan=2, pady=20)
        
        ttk.Button(button_frame, text="Save", command=self.save_settings).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.window.destroy).pack(side="left", padx=5)
        
        # Configure column weights
        self.window.columnconfigure(1, weight=1)
    
    def browse_folder(self):
        """Open folder browser dialog"""
        folder = filedialog.askdirectory(initialdir=self.save_folder_var.get())
        if folder:
            self.save_folder_var.set(folder)
    
    def save_settings(self):
        """Save settings and close dialog"""
        save_folder = self.save_folder_var.get().strip()
        companion_ip = self.companion_ip_var.get().strip()
        show_refresh = self.show_refresh_var.get()
        
        if not save_folder or not companion_ip:
            messagebox.showerror("Error", "Both Save Folder Path and Companion IP are required!")
            return
        
        self.config.update_settings(save_folder, companion_ip, show_refresh)
        self.refresh_callback()
        messagebox.showinfo("Success", "Settings saved successfully!")
        self.window.destroy()


class ServiceManagerGUI:
    """Main GUI application"""
    
    def __init__(self):
        # Setup logging
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
        
        # Initialize components
        self.config = Config()
        self.data_manager = None
        self.api = None
        
        # Setup GUI first
        self.setup_gui()
        
        # Then handle configuration and initialization
        self._handle_initial_setup()
    
    def _handle_initial_setup(self):
        """Handle initial setup after GUI is created"""
        # Check configuration
        if not self.config.is_valid():
            self._show_config_warning()
        
        self._initialize_components()
        
        # Check Companion connection at startup
        if self.api and not self._check_companion_connection():
            self._prompt_for_companion_ip()
        
        self._load_initial_data()
        """Check if Companion is reachable"""
        if not self.api:
            return False
        
        logging.info(f"Testing connection to Companion at {self.api.companion_ip}")
        return self.api.test_connection()
    
    def _check_companion_connection(self) -> bool:
        """Prompt user to update Companion IP if connection fails"""
        # Create a temporary root for the dialog
        temp_root = tk.Tk()
        temp_root.withdraw()
        
        dialog = CompanionConnectionDialog(
            temp_root, 
            self.config.get_companion_ip(), 
            self.config
        )
        
        new_ip = dialog.get_result()
        temp_root.destroy()
        
        if new_ip:
            # Reinitialize components with new IP
            self._initialize_components()
            logging.info(f"Updated Companion IP to: {new_ip}")
        else:
            logging.warning("Continuing with potentially unreachable Companion IP")
    
    def _show_config_warning(self):
        """Show configuration warning dialog"""
        messagebox.showwarning(
            "Configuration Missing", 
            "Save folder path or Companion IP is missing.\nPlease update the settings."
        )
        SettingsDialog(self.root, self.config, self._refresh_components)
    
    def _initialize_components(self):
        """Initialize data manager and API components"""
        save_folder = self.config.get_save_folder()
        companion_ip = self.config.get_companion_ip()
        
        if save_folder and companion_ip:
            json_path = Path(save_folder) / "selections.json"
            self.data_manager = ServiceDataManager(str(json_path))
            self.api = CompanionAPI(companion_ip)
    
    def _refresh_components(self):
        """Refresh components after settings change"""
        self._initialize_components()
        self._update_refresh_button_visibility()
        self.refresh_data()
        self.update_current_service_date()
    
    def setup_gui(self):
        """Setup the main GUI"""
        self.root = tk.Tk()
        self.root.title("Christ Church Service Activation")
        self.root.geometry("500x450")
        
        # Menu bar
        self._setup_menu()
        
        # Main content
        self._setup_main_content()
        
        # Update refresh button visibility
        self._update_refresh_button_visibility()
    
    def _setup_menu(self):
        """Setup menu bar"""
        menu_bar = tk.Menu(self.root)
        self.root.config(menu=menu_bar)
        
        settings_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Settings", menu=settings_menu)
        settings_menu.add_command(label="Open Settings", command=self.open_settings)
    
    def _setup_main_content(self):
        """Setup main content area"""
        # Title
        ttk.Label(
            self.root, 
            text="Activate Service Buttons", 
            font=("Arial", 16)
        ).pack(pady=(20, 10))
        
        # Current service label
        self.current_service_label = ttk.Label(
            self.root, 
            text="", 
            font=("Arial", 12)
        )
        self.current_service_label.pack(pady=(0, 10))
        
        # Date selection frame
        date_frame = ttk.Frame(self.root)
        date_frame.pack(pady=(0, 20), padx=20, fill='x')
        
        ttk.Label(date_frame, text="Select Date:", font=("Arial", 12)).pack(side='left', padx=10)
        
        self.date_var = tk.StringVar()
        self.date_dropdown = ttk.Combobox(
            date_frame, 
            textvariable=self.date_var, 
            state="readonly", 
            font=("Arial", 12), 
            width=15
        )
        self.date_dropdown.bind("<<ComboboxSelected>>", self.on_date_change)
        self.date_dropdown.pack(side='left', padx=10)
        
        # Song display
        self.song_vars = {
            key: tk.StringVar() 
            for key in ['song1', 'song2', 'song3', 'start', 'end', 'communion']
        }
        
        for key, var in self.song_vars.items():
            frame = ttk.Frame(self.root)
            frame.pack(fill='x', padx=20, pady=2)
            ttk.Label(frame, text=f"{key.capitalize()}:", font=("Arial", 12)).pack(side='left', padx=10)
            ttk.Label(frame, textvariable=var, font=("Arial", 12)).pack(side='right', padx=10)
        
        # Feedback label
        self.feedback_label = ttk.Label(
            self.root, 
            text="", 
            font=("Arial", 12), 
            foreground="green"
        )
        self.feedback_label.pack(pady=10)
        
        # Upload button
        self.upload_button = tk.Button(
            self.root, 
            text="Activate Service", 
            command=self.upload_data,
            width=20, 
            background="green", 
            foreground="white"
        )
        self.upload_button.pack(pady=10)
        
        # Refresh button
        self.refresh_button = ttk.Button(
            self.root, 
            text="Refresh Data", 
            command=self.refresh_data,
            width=20
        )
    
    def _update_refresh_button_visibility(self):
        """Update refresh button visibility based on settings"""
        if self.config.get_show_refresh_button():
            self.refresh_button.pack(pady=5)
        else:
            self.refresh_button.pack_forget()
    
    def _load_initial_data(self):
        """Load initial data and set default selections"""
        if not self.data_manager or not self.data_manager.data:
            self.feedback_label.config(
                text="❌ No service data found in selections.json.", 
                foreground="red"
            )
            return
        
        try:
            # Populate date dropdown
            date_options = self.data_manager.get_sorted_dates()
            self.date_dropdown['values'] = date_options
            
            # Set initial date
            nearest_service = self.data_manager.find_nearest_upcoming_service()
            if nearest_service:
                self.date_var.set(nearest_service[0])
                self.update_song_display(nearest_service[0])
            elif date_options:
                self.date_var.set(date_options[0])
                self.update_song_display(date_options[0])
                self.feedback_label.config(
                    text="⚠ No upcoming service found. Defaulted to earliest date.", 
                    foreground="orange"
                )
            
            self.update_current_service_date()
        except Exception as e:
            logging.error(f"Error loading initial data: {e}")
            self.feedback_label.config(
                text="❌ Error loading data", 
                foreground="red"
            )
    
    def on_date_change(self, event):
        """Handle date selection change"""
        selected_date = self.date_var.get()
        self.update_song_display(selected_date)
    
    def update_song_display(self, date_str: str):
        """Update song display for selected date"""
        if not self.data_manager:
            return
            
        service_data = self.data_manager.get_service_data(date_str)
        event_data = self.data_manager.data.get(date_str, {})
        
        for key, var in self.song_vars.items():
            value = service_data.get(key, "")
            path = event_data.get(f"{key}path", "")
            display_value = self._format_display_filename(value, path)
            var.set(display_value)
    
    def _format_display_filename(self, filename: str, path: str) -> str:
        """Format filename for display"""
        if not filename or not isinstance(filename, str):
            return "—"
        
        filename = filename.strip()
        filename_no_ext = os.path.splitext(filename)[0]
        file_extension = os.path.splitext(path)[1].lower()
        
        if file_extension == '.xspf':
            return f"{filename_no_ext} 🎵"
        return filename_no_ext
    
    def update_current_service_date(self):
        """Update current service date display"""
        if not self.api:
            self.current_service_label.config(text="API not configured")
            return
            
        current_date = self.api.get_current_service_date()
        if current_date:
            self.current_service_label.config(text=f"Current Service Date: {current_date}")
        else:
            self.current_service_label.config(text="Failed to fetch current service date")
    
    def upload_data(self):
        """Upload service data to Companion and launch VLC"""
        if not self.api or not self.data_manager:
            messagebox.showerror("Error", "API or data manager not configured")
            return
        
        self.upload_button.config(state="disabled", text="Activating...")
        
        try:
            selected_date = self.date_var.get()
            service_data = self.data_manager.get_service_data(selected_date)
            
            errors = self.api.update_service_data(service_data, selected_date)
            
            if errors:
                error_msg = "\n".join(errors[:5])  # Show first 5 errors
                messagebox.showwarning("Upload Warnings", f"Some updates failed:\n{error_msg}")
            
            self.feedback_label.config(
                text=f"Service for {selected_date} activated!", 
                foreground="green"
            )
            self.update_current_service_date()
            
            # Launch VLC
            self._launch_vlc()
            
        except Exception as e:
            logging.error(f"Upload failed: {e}")
            messagebox.showerror("Error", f"Upload failed: {e}")
            self.feedback_label.config(text="❌ Upload failed", foreground="red")
        
        finally:
            self.upload_button.config(state="normal", text="Activate Service")
    
    def _launch_vlc(self):
        """Launch VLC media player"""
        if VLCLauncher.launch_vlc():
            logging.info("VLC launched successfully")
            # Update feedback to include VLC launch
            current_text = self.feedback_label.cget("text")
            self.feedback_label.config(text=f"{current_text} VLC launched.")
        else:
            # Show manual launch prompt
            result = messagebox.askyesno(
                "VLC Launch Failed",
                "Could not automatically launch VLC.\n\n"
                "Would you like to launch VLC manually now?\n\n"
                "Click 'Yes' to open VLC installation folder\n"
                "Click 'No' to continue without VLC",
                icon='warning'
            )
            
            if result:
                self._open_vlc_folder()
    
    def _open_vlc_folder(self):
        """Open VLC installation folder in Explorer"""
        vlc_folders = [
            r"C:\Program Files\VideoLAN\VLC",
            r"C:\Program Files (x86)\VideoLAN\VLC",
        ]
        
        folder_opened = False
        for folder in vlc_folders:
            if os.path.exists(folder):
                try:
                    subprocess.run(["explorer", folder])
                    folder_opened = True
                    break
                except Exception as e:
                    logging.error(f"Failed to open folder {folder}: {e}")
        
        if not folder_opened:
            messagebox.showinfo(
                "VLC Not Found",
                "Could not locate VLC installation.\n\n"
                "Please install VLC Media Player from:\n"
                "https://www.videolan.org/vlc/"
            )
    
    def refresh_data(self):
        """Refresh service data from JSON"""
        if not self.data_manager:
            self.feedback_label.config(text="❌ Data manager not configured", foreground="red")
            return
        
        previous_date = self.date_var.get()
        
        if self.data_manager.load_data():
            # Update dropdown options
            date_options = self.data_manager.get_sorted_dates()
            self.date_dropdown['values'] = date_options
            
            # Restore previous selection if possible
            if previous_date in date_options:
                self.date_var.set(previous_date)
                self.update_song_display(previous_date)
            elif date_options:
                self.date_var.set(date_options[0])
                self.update_song_display(date_options[0])
            
            self.feedback_label.config(text="✅ Data refreshed successfully", foreground="blue")
        else:
            self.feedback_label.config(text="❌ Failed to refresh data", foreground="red")
    
    def open_settings(self):
        """Open settings dialog"""
        SettingsDialog(self.root, self.config, self._refresh_components)
    
    def run(self):
        """Start the GUI application"""
        self.root.mainloop()


def main():
    """Main entry point"""
    try:
        # Enable high DPI awareness on Windows
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass
        
        app = ServiceManagerGUI()
        app.run()
    except Exception as e:
        logging.error(f"Application error: {e}")
        # Create a simple error dialog even if the main app fails
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Fatal Error", f"Application failed to start: {e}")
        root.destroy()


if __name__ == "__main__":
    main()