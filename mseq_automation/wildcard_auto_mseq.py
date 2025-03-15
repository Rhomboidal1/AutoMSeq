# wildcard_auto_mseq.py
import os
import tkinter.filedialog as filedialog
from config import MseqConfig
from file_system_dao import FileSystemDAO
from folder_processor import FolderProcessor
from ui_automation import MseqAutomation
import re

def main():
    print("Starting wildcard auto mSeq...")
    
    import tkinter as tk
    root = tk.Tk()  # Create the root window
    root.withdraw()  # Hide it, but keep it as a parent for dialogs
    
    config = MseqConfig()
    print("Config loaded")
    file_dao = FileSystemDAO(config)
    print("FileSystemDAO initialized")
    ui_automation = MseqAutomation(config)
    print("UI Automation initialized")
    processor = FolderProcessor(file_dao, ui_automation, config)
    print("Folder processor initialized")
    
    # Get folder selection from user
    print("Asking for folder selection...")
    root.lift()  # Bring the dialog to the front
    root.attributes('-topmost', True)  # Keep it on top
    data_folder = filedialog.askdirectory(title="Select today's data folder to mseq folders")
    #data_folder = r"C:\Users\rhomb\Github\test_folder"  # Use a test folder that exists
    root.attributes('-topmost', False)
    if not data_folder:
        print("No folder selected, exiting")
        return
    print(f"Selected folder: {data_folder}")
    data_folder = re.sub(r'/', '\\\\', data_folder)
    
    # Get all folders without filtering
    print("Getting all folders...")
    all_folders = file_dao.get_folders(data_folder)
    print(f"Found {len(all_folders)} folders")
    
    # Process each folder
    for folder in all_folders:
        print(f"Processing folder: {folder}")
        processor.process_wildcard_folder(folder)
    
    # Close mSeq application
    print("Closing mSeq application...")
    ui_automation.close()
    print("All done!")

if __name__ == "__main__":
    main()