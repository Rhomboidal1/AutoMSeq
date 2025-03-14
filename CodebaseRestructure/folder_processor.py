# folder_processor.py
import re

class FolderProcessor:
    def __init__(self, file_dao, ui_automation, config):
        self.file_dao = file_dao
        self.ui_automation = ui_automation
        self.config = config
    
    def is_mseq_processed(self, folder):
        """Check if a folder has been processed by mSeq"""
        current_artifacts = []
        for item in os.listdir(folder):
            if item in self.config.MSEQ_ARTIFACTS:
                current_artifacts.append(item)
        return set(current_artifacts) == self.config.MSEQ_ARTIFACTS
    
    def has_output_files(self, folder):
        """Check if folder has all 5 expected output files"""
        count = 0
        for item in os.listdir(folder):
            if os.path.isfile(os.path.join(folder, item)):
                for extension in self.config.TEXT_FILES:
                    if item.endswith(extension):
                        count += 1
        return count == 5
    
    def process_bio_folder(self, folder):
        """Process a BioI folder (specialized for IND)"""
        # Implementation would use file_dao and ui_automation
        pass
    
    def process_plate_folder(self, folder):
        """Process a plate folder"""
        # Implementation for plate folders
        pass
    
    def process_wildcard_folder(self, folder):
        """Process any folder"""
        # Generic implementation
        pass