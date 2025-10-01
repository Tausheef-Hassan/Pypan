import os
import sys

# For PyInstaller bundled app
if getattr(sys, 'frozen', False):
    # Running as compiled executable
    # Use AppData for config files (persistent and writable)
    CONFIG_DIR = os.path.join(os.path.expanduser('~'), '.pypan')
    # Create config directory if it doesn't exist
    os.makedirs(CONFIG_DIR, exist_ok=True)
    # Pywikibot data files are in the bundled location
    PYWIKIBOT_DATA_DIR = sys._MEIPASS
else:
    # Running as script
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
    CONFIG_DIR = SCRIPT_DIR
    PYWIKIBOT_DATA_DIR = CONFIG_DIR

USER_CONFIG_PATH = os.path.join(CONFIG_DIR, 'user-config.py')
PASSWORD_FILE_PATH = os.path.join(CONFIG_DIR, 'user-password.py')

import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
import requests
import logging
from datetime import datetime
import shutil
from PIL import Image

ALLOWED_EXTENSIONS = {'.gif', '.jpg', '.jpeg', '.mid', '.midi', '.ogg', '.ogv', 
                      '.oga', '.png', '.svg', '.xcf', '.djvu', '.pdf', '.webm'}

def safe_chmod(path, mode):
    try:
        os.chmod(path, mode)
    except PermissionError:
        pass
    except Exception as e:
        print(f"Warning: could not chmod {path}: {e}")

class LoginWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Wikimedia Commons - Login")
        self.root.geometry("400x250")
        self.root.resizable(False, False)
        
        self.username = tk.StringVar()
        self.password = tk.StringVar()
        self.show_password = tk.BooleanVar(value=False)
        self.login_successful = False
        
        self.setup_ui()
        
        self.center_window()
        
    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        title_label = ttk.Label(main_frame, text="Wikimedia Commons Login", 
                               font=('Arial', 14, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        ttk.Label(main_frame, text="Username:").grid(row=1, column=0, sticky=tk.W, pady=5)
        username_entry = ttk.Entry(main_frame, textvariable=self.username, width=30)
        username_entry.grid(row=1, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        username_entry.focus()
        
        ttk.Label(main_frame, text="Password:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.password_entry = ttk.Entry(main_frame, textvariable=self.password, 
                                       show="●", width=30)
        self.password_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5)
        
        self.eye_button = ttk.Button(main_frame, text="👁", width=3, 
                                    command=self.toggle_password)
        self.eye_button.grid(row=2, column=2, sticky=tk.W, padx=(5, 0), pady=5)
        
        info_label = ttk.Label(main_frame, 
                              text="Enter your Wikimedia Commons credentials",
                              font=('Arial', 8), foreground='gray')
        info_label.grid(row=3, column=0, columnspan=3, pady=(10, 20))
        
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3)
        
        ttk.Button(button_frame, text="Login", command=self.login, 
                  width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Exit", command=self.exit_app, 
                  width=15).pack(side=tk.LEFT, padx=5)
        
        self.root.bind('<Return>', lambda e: self.login())
        
        main_frame.columnconfigure(1, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
    def toggle_password(self):
        if self.show_password.get():
            self.password_entry.config(show="●")
            self.show_password.set(False)
        else:
            self.password_entry.config(show="")
            self.show_password.set(True)
            
    def create_config_files(self, username, password):
        """Create Pywikibot configuration files"""
        try:
            # Creating user-config.py
            user_config_content = f"""# -*- coding: utf-8 -*-
family = 'commons'
mylang = 'commons'
usernames['commons']['commons'] = '{username}'
password_file = r"{PASSWORD_FILE_PATH}"
maxlag = 60
put_throttle = 1
console_encoding = 'utf-8'
max_retries = 10
simulate = False
textfile_encoding = 'utf-8'
"""
            with open(USER_CONFIG_PATH, 'w', encoding='utf-8') as f:
                f.write(user_config_content)
                f.flush()
                os.fsync(f.fileno())
            safe_chmod(USER_CONFIG_PATH, 0o600)

            # Creating user-password.py
            password_content = f"""# -*- coding: utf-8 -*-
('commons', 'commons', '{username}', '{password}')
"""
            with open(PASSWORD_FILE_PATH, 'w', encoding='utf-8') as f:
                f.write(password_content)
                f.flush()
                os.fsync(f.fileno())
            safe_chmod(PASSWORD_FILE_PATH, 0o600)
            
            time.sleep(0.5)

            return True

        except Exception as e:
            messagebox.showerror("Error", f"Failed to create config files:\n{str(e)}")
            return False
            
    def login(self):
        username = self.username.get().strip()
        password = self.password.get().strip()
        
        if not username or not password:
            messagebox.showerror("Error", "Please enter both username and password")
            return
            
        # Creating config files
        if self.create_config_files(username, password):
            self.login_successful = True
            self.root.destroy()
        
    def exit_app(self):
        self.root.destroy()
        sys.exit(0)


class PyPan:
    def __init__(self, root, username):
        self.root = root
        self.root.title("Wikimedia Commons Batch Uploader")
        self.root.geometry("800x600")
        
        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar(value="upload_results.xlsx")
        self.num_workers_value = 1
        self.username = username
        self.is_running = False
        self.is_paused = False
        self.total_files = 0
        self.processed_files = 0
        self.successful_uploads = 0
        self.failed_uploads = 0
        self.start_time = None
        self.executor = None
        self.site = None
        self.internet_status = tk.StringVar(value="Unknown")
        
        self.results = []
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[logging.StreamHandler()]
        )
        self.logger = logging.getLogger(__name__)
               
        self.setup_ui()
        
    def cleanup_config_files(self):
        """Delete config files and extra pywikibot artifacts in CONFIG_DIR"""
        try:
            # removing user-config and password files
            for p in (USER_CONFIG_PATH, PASSWORD_FILE_PATH):
                try:
                    if os.path.exists(p):
                        os.remove(p)
                except Exception as e:
                    print(f"Warning: could not remove {p}: {e}")

            # removing apicache directory
            apicache_dir = os.path.join(CONFIG_DIR, 'apicache')
            try:
                if os.path.isdir(apicache_dir):
                    shutil.rmtree(apicache_dir, ignore_errors=True)
            except Exception as e:
                print(f"Warning: could not remove apicache dir {apicache_dir}: {e}")

            # removing throttle control file
            throttle_path = os.path.join(CONFIG_DIR, 'throttle.ctrl')
            try:
                if os.path.exists(throttle_path):
                    os.remove(throttle_path)
            except Exception as e:
                print(f"Warning: could not remove {throttle_path}: {e}")

            # removing lwp file
            try:
                uname = getattr(self, 'username', '') or ''
                lwp_name = f"pywikibot-{uname.replace(' ', '_')}.lwp"
                lwp_path = os.path.join(CONFIG_DIR, lwp_name)
                if os.path.exists(lwp_path):
                    os.remove(lwp_path)
            except Exception as e:
                print(f"Warning: could not remove lwp file: {e}")

            # removing upload_log.txt
            try:
                log_path = os.path.join(CONFIG_DIR, 'upload_log.txt')
                if os.path.exists(log_path):
                    os.remove(log_path)
            except Exception as e:
                print(f"Warning: could not remove upload_log.txt: {e}")

        except Exception as e:
            print(f"Warning: cleanup_config_files failed: {e}")
        
    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        config_frame = ttk.LabelFrame(main_frame, text="Configuration", padding="10")
        config_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(config_frame, text="Logged in as:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        user_label = ttk.Label(config_frame, text=self.username, font=('Arial', 10, 'bold'))
        user_label.grid(row=0, column=1, sticky=tk.W, padx=(0, 10))
        
        ttk.Label(config_frame, text="Input Excel File:").grid(row=1, column=0, sticky=tk.W, padx=(0, 5))
        ttk.Entry(config_frame, textvariable=self.input_file, width=50).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 5))
        ttk.Button(config_frame, text="Browse", command=self.browse_input_file).grid(row=1, column=2)
        
        ttk.Label(config_frame, text="Output Excel File:").grid(row=2, column=0, sticky=tk.W, padx=(0, 5))
        ttk.Entry(config_frame, textvariable=self.output_file, width=30).grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Label(config_frame, text="Ignore Warnings:").grid(row=4, column=0, sticky=tk.W, padx=(0,5))
        self.ignore_warnings_var = tk.StringVar(value="True")
        ignore_dropdown = ttk.Combobox(config_frame, textvariable=self.ignore_warnings_var, values=["True", "False"], width=5, state="readonly")
        ignore_dropdown.grid(row=4, column=1, sticky=tk.W, padx=(0,10))
        
        ttk.Label(config_frame, text="Max Attempts:").grid(row=4, column=2, sticky=tk.W, padx=(0,5))
        self.max_attempts_var = tk.IntVar(value=10)
        ttk.Entry(config_frame, textvariable=self.max_attempts_var, width=5).grid(row=4, column=3, sticky=tk.W, padx=(0,10))

        ttk.Label(config_frame, text="Pause (s):").grid(row=4, column=4, sticky=tk.W, padx=(0,5))
        self.pause_seconds_var = tk.IntVar(value=10)
        ttk.Entry(config_frame, textvariable=self.pause_seconds_var, width=5).grid(row=4, column=5, sticky=tk.W, padx=(0,10))
        
        ttk.Label(config_frame, text="Internet Status:").grid(row=3, column=0, sticky=tk.W, padx=(0, 5))
        self.internet_status_label = ttk.Label(config_frame, textvariable=self.internet_status, foreground="gray")
        self.internet_status_label.grid(row=3, column=1, sticky=tk.W, padx=(0, 10))
        
        config_frame.columnconfigure(1, weight=1)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.start_btn = ttk.Button(button_frame, text="Start Upload", command=self.start_upload)
        self.start_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.pause_btn = ttk.Button(button_frame, text="Pause", command=self.pause_upload, state=tk.DISABLED)
        self.pause_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.stop_btn = ttk.Button(button_frame, text="Stop", command=self.stop_upload, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.test_connection_btn = ttk.Button(button_frame, text="Test Connection", command=lambda: self.test_internet_connection(show_success=True))
        self.test_connection_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding="10")
        progress_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        self.status_label = ttk.Label(progress_frame, text="Ready")
        self.status_label.grid(row=1, column=0, sticky=tk.W)
        
        self.stats_label = ttk.Label(progress_frame, text="Files: 0/0 | Success: 0 | Failed: 0")
        self.stats_label.grid(row=2, column=0, sticky=tk.W)
        
        self.time_label = ttk.Label(progress_frame, text="Time: 00:00:00 | ETA: --:--:--")
        self.time_label.grid(row=3, column=0, sticky=tk.W)
        
        progress_frame.columnconfigure(0, weight=1)
        
        log_frame = ttk.LabelFrame(main_frame, text="Log", padding="10")
        log_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        self.log_text = tk.Text(log_frame, height=15, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
    def browse_input_file(self):
        filename = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.input_file.set(filename)
            # Setting output file to same directory of input
            input_dir = os.path.dirname(filename)
            input_basename = os.path.basename(filename)
            input_name, input_ext = os.path.splitext(input_basename)
            output_filename = os.path.join(input_dir, f"{input_name}_results{input_ext}")
            self.output_file.set(output_filename)
            
    def log_message(self, message, level="INFO"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {level}: {message}\n"
        
        self.log_text.insert(tk.END, formatted_message)
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
        if level == "ERROR":
            self.logger.error(message)
        elif level == "WARNING":
            self.logger.warning(message)
        else:
            self.logger.info(message)
            
    def update_internet_status(self, status):
        """Update internet status in UI"""
        self.internet_status.set(status)
        if status == "Active":
            self.internet_status_label.config(foreground="green")
        elif status == "Inactive":
            self.internet_status_label.config(foreground="red")
        else:
            self.internet_status_label.config(foreground="gray")
            
    def test_internet_connection(self, show_success=False):
        """Test internet connectivity using multiple reliable endpoints"""
        test_urls = [
            "https://www.google.com",
            "https://commons.wikimedia.org",
            "https://www.cloudflare.com",
            "https://8.8.8.8" 
        ]
        
        for url in test_urls:
            try:
                response = requests.get(url, timeout=10)
                if response.status_code == 200:
                    self.update_internet_status("Active")
                    if show_success:
                        self.log_message("Internet connection is working")
                        messagebox.showinfo("Connection Test", "Internet connection is working!")
                    return True
            except:
                continue
                
        try:
            import socket
            socket.gethostbyname("google.com")
            self.update_internet_status("Active")
            return True
        except:
            pass
            
        self.update_internet_status("Inactive")
        if show_success:
            self.log_message("Internet connection failed", "ERROR")
        return False
            
    def wait_for_internet(self):
        """Wait for internet connection to be restored"""
        retry_count = 0
        while self.is_running and not self.test_internet_connection():
            retry_count += 1
            if retry_count == 1:  # Only log on first attempt
                self.log_message("Waiting for internet connection...")
            time.sleep(10)
        
        if self.is_running and retry_count > 0:
            self.log_message("Internet connection restored")
        return self.is_running
        
    def initialize_pywikibot(self):
        """Initialize Pywikibot with proper config paths"""
        try:
            # Set environment variables for pywikibot
            os.environ['PYWIKIBOT_DIR'] = CONFIG_DIR
        
            # Add CONFIG_DIR to sys.path so pywikibot can find user-config.py
            if CONFIG_DIR not in sys.path:
                sys.path.insert(0, CONFIG_DIR)
            
            # Force reload of pywikibot to use new config
            modules_to_remove = [key for key in sys.modules.keys() if key.startswith('pywikibot')]
            for module in modules_to_remove:
                del sys.modules[module]

            # Wait for config files to be readable
            for _ in range(10):  # up to ~1 second total (10 * 0.1s)
                if os.path.exists(USER_CONFIG_PATH) and os.path.exists(PASSWORD_FILE_PATH):
                    try:
                        # quick open test to ensure filesystem returns readable content
                        with open(PASSWORD_FILE_PATH, 'r', encoding='utf-8') as tf:
                            if tf.read(1) is not None:
                                break
                    except Exception:
                        pass
                time.sleep(0.1)
            else:
                self.log_message("Config files not ready for pywikibot.", "ERROR")
                return False
            
            import importlib
            pywikibot = importlib.import_module('pywikibot')
            from pywikibot import FilePage as _FilePage

            self.pywikibot = pywikibot
            self.FilePage = _FilePage

            # Check if config files exist
            if not os.path.exists(USER_CONFIG_PATH) or not os.path.exists(PASSWORD_FILE_PATH):
                self.log_message(f"Config files missing: {USER_CONFIG_PATH}, {PASSWORD_FILE_PATH}", "ERROR")
                return False
                
            # Verify password file is readable
            max_retries = 5
            for attempt in range(max_retries):
                try:
                    with open(PASSWORD_FILE_PATH, 'r', encoding='utf-8') as f:
                        content = f.read()
                        if content.strip():  # Ensure file has content
                            break
                except Exception as e:
                    if attempt < max_retries - 1:
                        time.sleep(0.5)
                    else:
                        self.log_message(f"Cannot read password file after {max_retries} attempts: {e}", "ERROR")
                        return False
            
            self.site = self.pywikibot.Site('commons', 'commons')
            self.site.login()

            # Debug info
            self.log_message(f"Config directory: {CONFIG_DIR}")
            if getattr(sys, 'frozen', False):
                self.log_message(f"Pywikibot data directory: {PYWIKIBOT_DATA_DIR}")
            self.log_message(f"Successfully logged in as {self.username}")
            return True

        except Exception as e:
            self.log_message(f"Failed to initialize Pywikibot: {str(e)}", "ERROR")
            import traceback
            self.log_message(f"Traceback: {traceback.format_exc()}", "ERROR")
            return False

    def get_extension_from_file(self, file_path):
        """Get extension from file, checking EXIF data if needed"""
        _, ext = os.path.splitext(file_path)
        
        if ext:
            return ext.lower()
        
        # No extension, try to get from EXIF/file type
        try:
            with Image.open(file_path) as img:
                format_ext = img.format.lower()
                if format_ext == 'jpeg':
                    return '.jpg'
                return f'.{format_ext}'
        except Exception as e:
            self.log_message(f"Could not determine file type from EXIF for {file_path}: {e}", "WARNING")
            return None
            
    def upload_single_file(self, row_data, row_index):
        """Upload a single file with retry logic"""
        from pywikibot.exceptions import UploadError
        file_path, target_filename, description = row_data
        
        # Check if file exists first
        if not os.path.exists(file_path):
            result = {
                'row': row_index + 1,
                'file_path': file_path,
                'target_filename': target_filename,
                'status': 'Skipped',
                'error': 'File not found',
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            return result
        
        # Get file extension
        file_ext = self.get_extension_from_file(file_path)
        
        # Skip if no extension could be determined
        if not file_ext:
            result = {
                'row': row_index + 1,
                'file_path': file_path,
                'target_filename': target_filename,
                'status': 'Skipped',
                'error': 'Could not determine file extension',
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            self.log_message(f"Skipping {file_path}: Could not determine file extension", "WARNING")
            return result
        
        # Check if extension is allowed
        if file_ext not in ALLOWED_EXTENSIONS:
            result = {
                'row': row_index + 1,
                'file_path': file_path,
                'target_filename': target_filename,
                'status': 'Skipped',
                'error': f'File extension {file_ext} not allowed',
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            self.log_message(f"Skipping {file_path}: Extension {file_ext} not in allowed list", "WARNING")
            return result
        
        # Add extension to target filename if not present
        if not os.path.splitext(target_filename)[1]:
            target_filename += file_ext
                
        result = {
            'row': row_index + 1,
            'file_path': file_path,
            'target_filename': target_filename,
            'status': 'Failed',
            'error': '',
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        for attempt in range(self.max_attempts_var.get()):
            try:
                # Check if paused
                while self.is_paused and self.is_running:
                    time.sleep(0.5)
                    
                if not self.is_running:
                    result['error'] = 'Upload stopped by user'
                    return result
                
                # Check internet connection before each upload
                if not self.test_internet_connection():
                    self.log_message(f"No internet connection for {target_filename}, waiting...", "WARNING")
                    if not self.wait_for_internet():
                        result['error'] = 'No internet connection'
                        return result
                
                # Check if file exists
                if not os.path.exists(file_path):
                    result['error'] = f'File not found: {file_path}'
                    return result
                
                # Create FilePage
                file_page = self.FilePage(self.site, f'File:{target_filename}')
                
                # Check if file already exists
                if file_page.exists():
                    result['status'] = 'Skipped'
                    result['error'] = 'File already exists'
                    return result
                
                # Upload file
                self.log_message(f"Uploading {target_filename} (attempt {attempt + 1})")
                
                success = file_page.upload(
                    source=file_path,
                    comment=f"Pypan 0.1.1a0",
                    text=description,
                    ignore_warnings=(self.ignore_warnings_var.get() == "True")
                )
                
                if success:
                    result['status'] = 'Success'
                    result['error'] = ''
                    self.log_message(f"Successfully uploaded {target_filename}")
                    
                    # Wait 0.2 seconds after successful upload
                    time.sleep(0.2)
                    return result
                else:
                    result['error'] = 'Upload failed - server response'
                    
            except UploadError as e:
                result['error'] = f'Upload warning: {str(e)}'
                self.log_message(f"Upload warning for {target_filename}: {str(e)}", "WARNING")
                
            except Exception as e:
                result['error'] = f'Exception: {str(e)}'
                self.log_message(f"Error uploading {target_filename}: {str(e)}", "ERROR")
            
            # Wait before retry 
            if attempt < self.max_attempts_var.get() - 1:
                self.log_message(f"Waiting {self.pause_seconds_var.get()} seconds before retry (attempt {attempt + 1}/{self.max_attempts_var.get()})")
                time.sleep(self.pause_seconds_var.get())
                
        return result
        
    def update_progress(self):
        """Update progress indicators"""
        if self.total_files > 0:
            progress = (self.processed_files / self.total_files) * 100
            self.progress_var.set(progress)
            
            # Update stats
            self.stats_label.config(
                text=f"Files: {self.processed_files}/{self.total_files} | "
                     f"Success: {self.successful_uploads} | "
                     f"Failed: {self.failed_uploads}"
            )
            
            # Calculate time and ETA
            if self.start_time:
                elapsed = time.time() - self.start_time
                elapsed_str = time.strftime("%H:%M:%S", time.gmtime(elapsed))
                
                if self.processed_files > 0:
                    avg_time_per_file = elapsed / self.processed_files
                    remaining_files = self.total_files - self.processed_files
                    eta_seconds = avg_time_per_file * remaining_files
                    eta_str = time.strftime("%H:%M:%S", time.gmtime(eta_seconds))
                else:
                    eta_str = "--:--:--"
                    
                self.time_label.config(text=f"Time: {elapsed_str} | ETA: {eta_str}")
                
    def save_results(self):
        """Save results to Excel file in same directory as input with status in last column"""
        try:
            # Reading
            input_df = pd.read_excel(self.input_file.get(), header=None)
            
            # Creating column
            input_df['Upload_Status'] = ''
            
            # Update status for each row based on results
            for result in self.results:
                row_idx = result['row'] - 1  # Convert to 0-based index
                if row_idx < len(input_df):
                    if result['status'] == 'Success':
                        input_df.loc[row_idx, 'Upload_Status'] = 'Success'
                    elif result['status'] == 'Skipped':
                        input_df.loc[row_idx, 'Upload_Status'] = f"Skipped: {result['error']}"
                    else:
                        input_df.loc[row_idx, 'Upload_Status'] = f"Failed: {result['error']}"
            
            # Save to output file
            input_df.to_excel(self.output_file.get(), index=False, header=False)
            self.log_message(f"Results saved to {self.output_file.get()}")
        except Exception as e:
            self.log_message(f"Error saving results: {str(e)}", "ERROR")
            
    def upload_worker_thread(self):
        """Main upload worker thread"""
        try:
            df = pd.read_excel(self.input_file.get(), header=None)
            self.total_files = len(df)
            
            if self.total_files == 0:
                self.log_message("No files found in Excel file", "ERROR")
                return
                
            self.log_message(f"Found {self.total_files} files to upload")
            
            if not self.initialize_pywikibot():
                return
                
            # Process files with thread pool
            with ThreadPoolExecutor(max_workers=self.num_workers_value) as executor:
                self.executor = executor
                
                # Submit all tasks
                future_to_row = {}
                for index, row in df.iterrows():
                    if not self.is_running:
                        break
                        
                    file_path = str(row[0]) if pd.notna(row[0]) else ""
                    target_filename = str(row[1]) if pd.notna(row[1]) else ""
                    description = str(row[2]) if pd.notna(row[2]) else ""
                    description += "\n[[Category: Uploaded with pypan]]"
                    
                    if file_path and target_filename:
                        future = executor.submit(
                            self.upload_single_file, 
                            (file_path, target_filename, description),
                            index
                        )
                        future_to_row[future] = index
                        
                # Process completed tasks
                for future in as_completed(future_to_row):
                    if not self.is_running:
                        break
                        
                    try:
                        result = future.result()
                        self.results.append(result)
                        
                        self.processed_files += 1
                        
                        if result['status'] == 'Success':
                            self.successful_uploads += 1
                        else:
                            self.failed_uploads += 1
                            
                        self.update_progress()
                        
                    except Exception as e:
                        self.log_message(f"Error processing result: {str(e)}", "ERROR")
                        self.failed_uploads += 1
                        
            self.save_results()
            
        except Exception as e:
            self.log_message(f"Upload thread error: {str(e)}", "ERROR")
        finally:
            self.upload_finished()
            
    def upload_finished(self):
        """Called when upload process is finished"""
        self.is_running = False
        self.is_paused = False
        
        self.start_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.DISABLED, text="Pause")
        self.stop_btn.config(state=tk.DISABLED)
        
        self.status_label.config(text="Upload completed")
        
        self.log_message("Upload process completed")
        self.log_message(f"Total: {self.total_files}, Success: {self.successful_uploads}, Failed: {self.failed_uploads}")
        
        # completion dialog
        messagebox.showinfo(
            "Upload Complete",
            f"Batch upload completed!\n\n"
            f"Total files: {self.total_files}\n"
            f"Successful: {self.successful_uploads}\n"
            f"Failed: {self.failed_uploads}\n\n"
            f"Results saved to:\n{self.output_file.get()}"
        )
        
        # Cleanup config files
        self.cleanup_config_files()
        
    def start_upload(self):
        """Start the upload process"""
        if not self.input_file.get():
            messagebox.showerror("Error", "Please select an input file")
            return
            
        if not os.path.exists(self.input_file.get()):
            messagebox.showerror("Error", "Input file does not exist")
            return
            
        # Reset counters
        self.processed_files = 0
        self.successful_uploads = 0
        self.failed_uploads = 0
        self.results = []
        self.start_time = time.time()
        
        self.is_running = True
        self.is_paused = False
        self.start_btn.config(state=tk.DISABLED)
        self.pause_btn.config(state=tk.NORMAL, text="Pause")
        self.stop_btn.config(state=tk.NORMAL)
        
        self.status_label.config(text="Starting upload...")
        self.log_message("Starting upload process")
        
        # Start worker thread
        threading.Thread(target=self.upload_worker_thread, daemon=True).start()
        
    def pause_upload(self):
        """Pause/resume the upload process"""
        if self.is_paused:
            self.is_paused = False
            self.pause_btn.config(text="Pause")
            self.status_label.config(text="Resuming upload...")
            self.log_message("Upload resumed")
        else:
            self.is_paused = True
            self.pause_btn.config(text="Resume")
            self.status_label.config(text="Upload paused")
            self.log_message("Upload paused")
            
    def stop_upload(self):
        """Stop the upload process"""
        self.is_running = False
        self.is_paused = False
        
        if self.executor:
            self.executor.shutdown(wait=False)
            
        self.status_label.config(text="Stopping upload...")
        self.log_message("Upload stopped by user")

def main():
    # Create config directory if running as exe
    if getattr(sys, 'frozen', False):
        os.makedirs(CONFIG_DIR, exist_ok=True)
    
    login_root = tk.Tk()
    login_app = LoginWindow(login_root)
    login_root.mainloop()
    
    if login_app.login_successful:
        root = tk.Tk()
        app = PyPan(root, login_app.username.get())
        root.mainloop()
    else:
        print("Login cancelled or failed. Exiting.")


if __name__ == "__main__":
    main()