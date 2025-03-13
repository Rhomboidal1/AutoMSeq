import os
import shutil
import re
import logging
import time
from datetime import datetime, timedelta

# Set up logging
logging.basicConfig(
    filename='file_copy.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Define the source and destination directories
source_directory = 'C:/P/Data/Individuals/'
destination_directory = 'C:/P/Data/Controls/'

# Define the list of allowed filename patterns (regular expressions)
allowed_filename_patterns = [
    r'^_Water.*\.txt$',# Files starting with "_Water" and ending with ".txt"
    r'^_pGEM.*\.txt$',# Files starting with "_pGEM" followed by 4 digits and ".txt"
]

# Get the current time and calculate the cutoff time (30 days ago)
current_time = time.time()
days_ago_30 = current_time - (30 * 24 * 60 * 60)

def is_file_allowed(filename, patterns):
    """Check if a filename matches any of the allowed patterns."""
    for pattern in patterns:
        if re.match(pattern, filename):
            return True
    return False

def is_recent_folder(path):
    """Check if the folder was created or modified within the last 30 days."""
    folder_time = os.path.getmtime(path)  # You can use getctime() for creation time instead
    return folder_time > days_ago_30

def copy_and_rename_files(src_dir, dest_dir, allowed_patterns):
    # Walk through all the directories and files in the source directory
    for root, dirs, files in os.walk(src_dir):
        # Only process folders that have been created/modified in the last 30 days
        if is_recent_folder(root):
            for file in files:
                # Check if the file matches any of the allowed regular expression patterns
                if is_file_allowed(file, allowed_patterns):
                    # Construct full file path
                    full_file_path = os.path.join(root, file)

                    # Create a prefix based on the current directory relative to the source directory
                    relative_dir = os.path.relpath(root, src_dir)
                    prefix = relative_dir.replace(os.sep, '_')

                    # Create the new file name with the directory prefix
                    new_file_name = f'{prefix}_{file}'

                    # Construct the destination path
                    dest_file_path = os.path.join(dest_dir, new_file_name)

                    try:
                        # Copy the file to the destination with the new name
                        shutil.copy2(full_file_path, dest_file_path)
                        logging.info(f'Copied: {full_file_path} to {dest_file_path}')
                        print(f'Copied: {full_file_path} to {dest_file_path}')
                    except Exception as e:
                        logging.error(f'Error copying {full_file_path}: {e}')
                        print(f'Error copying {full_file_path}: {e}')
                else:
                    logging.info(f'Skipped: {file} (does not match any pattern)')
                    print(f'Skipped: {file} (does not match any pattern)')
        else:
            logging.info(f'Skipped folder: {root} (not modified/created in the last 30 days)')
            print(f'Skipped folder: {root} (not modified/created in the last 30 days)')

# Make sure the destination directory exists, create if not
if not os.path.exists(destination_directory):
    os.makedirs(destination_directory)
    logging.info(f'Created destination directory: {destination_directory}')

# Call the function to process the files
copy_and_rename_files(source_directory, destination_directory, allowed_filename_patterns)