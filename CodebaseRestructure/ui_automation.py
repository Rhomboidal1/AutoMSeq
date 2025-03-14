# ui_automation.py
from pywinauto import Application, timings
from pywinauto.keyboard import send_keys

class MseqAutomation:
    def __init__(self, config):
        self.config = config
        self.app = None
        self.first_time_browsing = True
    
    def connect_or_start_mseq(self):
        """Connect to existing mSeq or start a new instance"""
        # Implementation from GetMseq function
        pass
    
    def process_folder(self, folder_path):
        """Run mSeq processing on a folder"""
        # Implementation from MseqOrder function
        pass
    
    def navigate_folder_tree(self, dialog, path):
        """Navigate folder tree in file dialog"""
        # Implementation from NavigateToFolder function
        pass
    
    def wait_for_dialog(self, dialog_type, timeout=None):
        """Wait for a specific dialog to appear"""
        # Standardized waiting function
        pass