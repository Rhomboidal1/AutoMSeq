# ui_automation.py
import os
import time
from pywinauto import Application, timings
from pywinauto.keyboard import send_keys
from pywinauto.findwindows import ElementNotFoundError, ElementAmbiguousError

class MseqAutomation:
    def __init__(self, config):
        self.config = config
        self.app = None
        self.main_window = None
        self.first_time_browsing = True
    
    def connect_or_start_mseq(self):
        """Connect to existing mSeq or start a new instance"""
        try:
            self.app = Application(backend='win32').connect(title_re='Mseq*', timeout=1)
        except (ElementNotFoundError, timings.TimeoutError):
            try:
                self.app = Application(backend='win32').connect(title_re='mSeq*', timeout=1)
            except (ElementNotFoundError, timings.TimeoutError):
                start_cmd = f'cmd /c "cd /d {self.config.MSEQ_PATH} && {self.config.MSEQ_EXECUTABLE}"'
                self.app = Application(backend='win32').start(start_cmd, wait_for_idle=False) 
                self.app.connect(title='mSeq', timeout=10)
            except ElementAmbiguousError:
                self.app = Application(backend='win32').connect(title_re='mSeq*', found_index=0, timeout=1)
                app2 = Application(backend='win32').connect(title_re='mSeq*', found_index=1, timeout=1)
                app2.kill()
        except ElementAmbiguousError:
            self.app = Application(backend='win32').connect(title_re='Mseq*', found_index=0, timeout=1)
            app2 = Application(backend='win32').connect(title_re='Mseq*', found_index=1, timeout=1)
            app2.kill()
        
        # Get the main window
        if not self.app.window(title_re='mSeq*').exists():
            self.main_window = self.app.window(title_re='Mseq*')
        else:
            self.main_window = self.app.window(title_re='mSeq*')
            
        return self.app, self.main_window
    
    def wait_for_dialog(self, dialog_type):
        """Wait for a specific dialog to appear"""
        timeout = self.config.TIMEOUTS.get(dialog_type, 5)
        
        if dialog_type == "browse_dialog":
            return timings.wait_until(timeout=timeout, retry_interval=0.1, 
                                     func=lambda: self.app.window(title='Browse For Folder').exists(), 
                                     value=True)
        elif dialog_type == "preferences":
            return timings.wait_until(timeout=timeout, retry_interval=0.1, 
                                     func=lambda: self.app.window(title='Mseq Preferences').exists(), 
                                     value=True)
        elif dialog_type == "copy_files":
            return timings.wait_until(timeout=timeout, retry_interval=0.1,
                                     func=lambda: self.app.window(title='Copy sequence files').exists(),
                                     value=True)
        elif dialog_type == "error_window":
            return timings.wait_until(timeout=timeout, retry_interval=0.5,
                                     func=lambda: self.app.window(title='File error').exists(),
                                     value=True)
        elif dialog_type == "call_bases":
            return timings.wait_until(timeout=timeout, retry_interval=0.3,
                                     func=lambda: self.app.window(title='Call bases?').exists(),
                                     value=True)
        elif dialog_type == "read_info":
            return timings.wait_until(timeout=timeout, retry_interval=0.1,
                                     func=lambda: self.app.window(title_re='Read information for*').exists(),
                                     value=True)
    
    def is_process_complete(self, folder_path):
        """Check if mSeq has finished processing the folder"""
        if self.app.window(title="Low quality files skipped").exists():
            return True
        
        # Check if all 5 text files have been created
        count = 0
        for item in os.listdir(folder_path):
            if os.path.isfile(os.path.join(folder_path, item)):
                for extension in self.config.TEXT_FILES:
                    if item.endswith(extension):
                        count += 1
        return count == 5
    
    def navigate_folder_tree(self, dialog, path):
        """Navigate the folder tree in a file dialog"""
        dialog.set_focus()
        tree_view = dialog.child_window(class_name="SysTreeView32")
        
        # Get the virtual folder name
        import win32com.client
        shell = win32com.client.Dispatch("Shell.Application")
        namespace = shell.Namespace(0x11)  # CSIDL_DRIVES
        virtual_folder = namespace.Title
        
        # Parse the path
        if ":" in path:
            # Path has a drive letter
            parts = path.split("\\")
            drive = parts[0]  # e.g., "P:"
            folders = parts[1:] if len(parts) > 1 else []
        else:
            # Network path
            parts = path.split("\\")
            drive = "\\" + "\\".join(parts[:3])  # e.g., \\server\share
            folders = parts[3:] if len(parts) > 3 else []
        
        try:
            # Start from Desktop
            desktop_item = tree_view.get_item('\\Desktop')
            desktop_item.expand()
            time.sleep(0.3)
            
            # Navigate to This PC
            this_pc_item = None
            for child in desktop_item.children():
                if "PC" in child.text() or "Computer" in child.text() or virtual_folder in child.text():
                    this_pc_item = child
                    break
            
            if not this_pc_item:
                raise ValueError(f"Could not find This PC or Computer in tree view")
            
            this_pc_item.expand()
            time.sleep(0.3)
            
            # Navigate to the drive
            drive_found = False
            mapped_name = self.config.NETWORK_DRIVES.get(drive, None)
            
            for drive_item in this_pc_item.children():
                drive_text = drive_item.text()
                
                # Check for both exact match and mapped name
                if (drive == drive_text or 
                    drive in drive_text or 
                    (mapped_name and mapped_name in drive_text)):
                    
                    dialog.set_focus()
                    drive_item.click_input()
                    drive_found = True
                    current_item = drive_item
                    time.sleep(0.3)
                    break
            
            if not drive_found:
                raise ValueError(f"Could not find drive '{drive}' or '{mapped_name}' in tree view")
            
            # Navigate through subfolders
            for folder in folders:
                current_item.expand()
                time.sleep(0.3)
                
                # Look for exact match first
                folder_found = False
                for child in current_item.children():
                    if child.text() == folder:
                        dialog.set_focus()
                        child.click_input()
                        folder_found = True
                        current_item = child
                        time.sleep(0.3)
                        break
                
                if not folder_found:
                    # Try partial match
                    for child in current_item.children():
                        if folder.lower() in child.text().lower():
                            dialog.set_focus()
                            child.click_input()
                            folder_found = True
                            current_item = child
                            time.sleep(0.3)
                            break
                
                if not folder_found:
                    raise ValueError(f"Could not find folder '{folder}' or similar")
            
            # Final folder should now be selected
            return True
        
        except Exception as e:
            # Log the error but don't raise it - we want to continue even if navigation fails
            print(f"Error during folder navigation: {e}")
            return False
    
    def process_folder(self, folder_path):
        """Process a folder with mSeq"""
        try:
            if not os.path.exists(folder_path):
                print(f"Warning: Folder does not exist: {folder_path}")
                return False
                
            # Check if there are any .ab1 files to process
            ab1_files = [f for f in os.listdir(folder_path) if f.endswith('.ab1')]
            if not ab1_files:
                print(f"No .ab1 files found in {folder_path}, skipping processing")
                return False
        except Exception as e:
            print(f"Error checking folder {folder_path}: {e}")
            return False
        
        self.app, self.main_window = self.connect_or_start_mseq()
        self.main_window.set_focus()
        send_keys('^n')  # Ctrl+N for new project
        
        # Wait for and handle Browse For Folder dialog
        self.wait_for_dialog("browse_dialog")
        dialog_window = self.app.window(title='Browse For Folder')
        
        # Add a delay for the first browsing operation
        if self.first_time_browsing:
            self.first_time_browsing = False
            time.sleep(0.5)  # mSeq needs time to initialize file browsing
        else:
            time.sleep(0.3)
        
        # Navigate to the target folder
        self.navigate_folder_tree(dialog_window, folder_path)
        ok_button = self.app.BrowseForFolder.child_window(title="OK", class_name="Button")
        ok_button.click_input()
        
        # Handle mSeq Preferences dialog
        self.wait_for_dialog("preferences")
        pref_window = self.app.window(title='Mseq Preferences')
        ok_button = pref_window.child_window(title="&OK", class_name="Button")
        ok_button.click_input()
        
        # Handle Copy sequence files dialog
        self.wait_for_dialog("copy_files")
        copy_files_window = self.app.window(title='Copy sequence files')
        shell_view = copy_files_window.child_window(title="ShellView", class_name="SHELLDLL_DefView")
        list_view = shell_view.child_window(class_name="DirectUIHWND")
        list_view.click_input()
        send_keys('^a')  # Select all files
        open_button = copy_files_window.child_window(title="&Open", class_name="Button")
        open_button.click_input()
        
        # Handle File error dialog
        self.wait_for_dialog("error_window")
        error_window = self.app.window(title='File error')
        error_ok_button = error_window.child_window(class_name="Button")
        error_ok_button.click_input()
        
        # Handle Call bases dialog
        self.wait_for_dialog("call_bases")
        call_bases_window = self.app.window(title='Call bases?')
        yes_button = call_bases_window.child_window(title="&Yes", class_name="Button")
        yes_button.click_input()
        
    # Wait for processing to complete with graceful timeout handling
        try:
            timings.wait_until(
                timeout=self.config.TIMEOUTS["process_completion"],
                retry_interval=0.2,
                func=lambda: self.is_process_complete(folder_path),
                value=True
            )
        except timings.TimeoutError:
            print(f"Warning: Timeout waiting for processing to complete for {folder_path}")
            print("This may be normal if the folder has already been processed or has special files")
        
        # Handle Low quality files skipped dialog if it appears
        if self.app.window(title="Low quality files skipped").exists():
            low_quality_window = self.app.window(title='Low quality files skipped')
            ok_button = low_quality_window.child_window(class_name="Button")
            ok_button.click_input()
        
        # Handle Read information dialog
        try:
            self.wait_for_dialog("read_info")
            if self.app.window(title_re='Read information for*').exists():
                read_window = self.app.window(title_re='Read information for*')
                read_window.close()
        except timings.TimeoutError:
            print(f"Read information dialog did not appear for {folder_path}")
            # Continue processing
        
        return True
    
    def close(self):
        """Close the mSeq application"""
        if self.app:
            self.app.kill()