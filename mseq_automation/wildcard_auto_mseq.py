# wildcard_auto_mseq.py
import os
import tkinter.filedialog as filedialog
from config import MseqConfig
from file_system_dao import FileSystemDAO
from folder_processor import FolderProcessor
from ui_automation import MseqAutomation
import re

def main():
    config = MseqConfig()
    file_dao = FileSystemDAO(config)
    ui_automation = MseqAutomation(config)
    processor = FolderProcessor(file_dao, ui_automation, config)
    
    # Get folder selection from user
    data_folder = filedialog.askdirectory(title="Select today's data folder to mseq folders")
    data_folder = re.sub(r'/', '\\\\', data_folder)
    
    # Get all folders without filtering
    all_folders = file_dao.get_folders(data_folder)
    
    # Process each folder
    for folder in all_folders:
        print(folder)
        processor.process_wildcard_folder(folder)
    
    # Close mSeq application
    ui_automation.close()

if __name__ == "__main__":
    main()