# file_system_dao.py
import os
import re
from shutil import move

class FileSystemDAO:
    def __init__(self, config):
        self.config = config
    
    def get_folders(self, path, pattern=None):
        """Get folders matching an optional regex pattern"""
        folders = []
        for item in os.listdir(path):
            full_path = os.path.join(path, item)
            if os.path.isdir(full_path):
                if pattern is None or re.search(pattern, item.lower()):
                    folders.append(full_path)
        return folders
    
    def get_files_by_extension(self, folder, extension):
        """Get all files with specified extension in a folder"""
        files = []
        for item in os.listdir(folder):
            if item.endswith(extension):
                files.append(os.path.join(folder, item))
        return files
    
    def contains_file_type(self, folder, extension):
        """Check if folder contains files with specified extension"""
        for item in os.listdir(folder):
            if item.endswith(extension):
                return True
        return False
    
    def create_folder_if_not_exists(self, path):
        """Create folder if it doesn't exist"""
        if not os.path.exists(path):
            os.mkdir(path)
        return path
    
    def move_folder(self, source, destination):
        """Move folder with proper error handling"""
        try:
            move(source, destination)
            return True
        except Exception as e:
            # Proper logging would be implemented here
            print(f"Error moving folder {source}: {e}")
            return False