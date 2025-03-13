import os
import shutil

# Define the list of allowed filenames (without the directory part)
allowed_filenames = [
    '_Water_M13F20.txt', '_Water_M13R27.txt', '_Water_T7Promoter.txt', '_Water_SP6Promoter.txt',
    '_pGEM_M13F20.txt', '_pGEM_M13R27.txt', '_pGEM_T7Promoter.txt', '_pGEM_SP6Promoter.txt'
]

# Define the source and destination directories
source_directory = 'C:/P/Data/Individuals/'
destination_directory = 'C:/P/Data/Controls/'

def copy_and_rename_files(src_dir, dest_dir, allowed_files):
    # Walk through all the directories and files in the source directory
    for root, dirs, files in os.walk(src_dir):
        for file in files:
            # Only process files that are in the list of allowed filenames
            if file in allowed_files:
                # Construct full file path
                full_file_path = os.path.join(root, file)

                # Create a prefix based on the current directory relative to the source directory
                relative_dir = os.path.relpath(root, src_dir)
                prefix = relative_dir.replace(os.sep, '_')

                # Create the new file name with the directory prefix
                new_file_name = f'{prefix}_{file}'

                # Construct the destination path
                dest_file_path = os.path.join(dest_dir, new_file_name)

                # Copy the file to the destination with the new name
                shutil.copy2(full_file_path, dest_file_path)
                print(f'Copied: {full_file_path} to {dest_file_path}')
            else:
                print(f'Skipped: {file} (not in allowed list)')

# Make sure the destination directory exists, create if not
if not os.path.exists(destination_directory):
    os.makedirs(destination_directory)

# Call the function to process the files
copy_and_rename_files(source_directory, destination_directory, allowed_filenames)
