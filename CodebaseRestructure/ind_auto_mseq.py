# ind_auto_mseq.py
from config import MseqConfig
from file_system_dao import FileSystemDAO
from folder_processor import FolderProcessor
from ui_automation import MseqAutomation
import sys

# Check for 32-bit Python requirement
if sys.maxsize > 2**32:
    # Relaunch in 32-bit logic
    pass

def main():
    config = MseqConfig()
    file_dao = FileSystemDAO(config)
    ui_automation = MseqAutomation(config)
    processor = FolderProcessor(file_dao, ui_automation, config)
    
    # Get folder selection from user
    data_folder = select_folder("Select today's data folder to mseq orders")
    
    # Process BioI folders
    bio_folders = file_dao.get_folders(data_folder, r'bioi-\d+')
    immediate_orders = file_dao.get_folders(data_folder, r'bioi-\d+_.+_\d+')
    pcr_folders = file_dao.get_folders(data_folder, r'fb-pcr\d+_\d+')
    
    for folder in bio_folders:
        processor.process_bio_folder(folder)
    
    for folder in immediate_orders:
        processor.process_order_folder(folder)
    
    for folder in pcr_folders:
        processor.process_pcr_folder(folder)
    
    # Close mSeq application
    ui_automation.close()

if __name__ == "__main__":
    main()