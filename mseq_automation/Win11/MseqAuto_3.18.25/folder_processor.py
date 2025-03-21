# folder_processor.py
import os
import re
import numpy as np
from datetime import datetime

class FolderProcessor:
    def __init__(self, file_dao, ui_automation, config, logger=None):
        self.file_dao = file_dao
        self.ui_automation = ui_automation
        self.config = config
        self.logger = logger 
        self.order_key_index = None  # Will be populated when needed

    def build_order_key_index(self, order_key):
        """Build lookup index for faster order key searches"""
        if self.order_key_index is not None:
            return  # Already built
            
        self.order_key_index = {}
        
        # Process each entry in the order key
        for entry in order_key:
            i_num, acct_name, order_num, sample_name = entry[0:4]
            normalized_name = self.file_dao.normalize_filename(sample_name, remove_extension=False)
            
            # Create entry in index (handle multiple entries with same normalized name)
            if normalized_name not in self.order_key_index:
                self.order_key_index[normalized_name] = []
            self.order_key_index[normalized_name].append((i_num, acct_name, order_num))
        
        self.logger(f"Built order key index with {len(self.order_key_index)} unique entries")

    def sort_customer_file(self, file_path, order_key, recent_inumbers):
        """Sort a customer file based on order key using the index"""
        # Build index if not already done
        if self.order_key_index is None:
            self.build_order_key_index(order_key)
            
        file_name = os.path.basename(file_path)
        # Only log once, not for each transformation step
        self.logger(f"Processing customer file: {file_name}")
        
        # Normalize the filename for matching
        normalized_name = self.file_dao.normalize_filename(file_name)
        
        # Check if we have this filename in our index
        if normalized_name in self.order_key_index:
            matches = self.order_key_index[normalized_name]
            
            # Prioritize matches from current folder's I number
            current_i_num = self.file_dao.get_inumber_from_name(os.path.dirname(file_path))
            
            for i_num, acct_name, order_num in matches:
                # Prioritize current I number if available
                if current_i_num and i_num == current_i_num:
                    destination_folder = self._create_and_get_order_folder(i_num, acct_name, order_num)
                    return self._move_file_to_destination(file_path, destination_folder, normalized_name)
            
            # If no match with current I number, use the first match
            i_num, acct_name, order_num = matches[0]
            destination_folder = self._create_and_get_order_folder(i_num, acct_name, order_num)
            return self._move_file_to_destination(file_path, destination_folder, normalized_name)
                
        # No match found
        self.logger(f"No match found in order key for: {normalized_name}")
        return False

    def _create_and_get_order_folder(self, i_num, acct_name, order_num):
        """Create order folder structure and return the path"""
        # Create order folder name 
        order_folder_name = f"BioI-{i_num}_{acct_name}_{order_num}"
        self.logger(f"Target order folder: {order_folder_name}")
        
        # Get parent folder path
        parent_folder = self._get_destination_for_order_by_inum(i_num)
        self.logger(f"Parent folder: {parent_folder}")
        
        # Create BioI folder if it doesn't exist
        bioi_folder_name = f"BioI-{i_num}"
        bioi_folder_path = os.path.join(parent_folder, bioi_folder_name)
        if not os.path.exists(bioi_folder_path):
            os.makedirs(bioi_folder_path)
            self.logger(f"Created BioI folder: {bioi_folder_path}")
        
        # Create full order folder path inside BioI folder
        order_folder_path = os.path.join(bioi_folder_path, order_folder_name)
        self.logger(f"Order folder path: {order_folder_path}")
        
        # Create order folder if it doesn't exist
        if not os.path.exists(order_folder_path):
            os.makedirs(order_folder_path)
            self.logger(f"Created order folder: {order_folder_path}")
            
        return order_folder_path

    def get_destination_for_order(self, order_folder, base_path):
        """Determine the correct destination for an order folder"""
        folder_name = os.path.basename(order_folder)
        self.logger(f"Processing order folder: {folder_name}")
        
        # Normalize the folder name
        normalized_folder_name = self.file_dao.normalize_filename(folder_name, remove_extension=False)
        
        i_num = self.file_dao.get_inumber_from_name(normalized_folder_name)
        self.logger(f"Extracted I number: {i_num}")
    
        if i_num:
            # Navigate up one level to get to the day folder
            day_data_path = os.path.dirname(base_path)
            self.logger(f"Searching for matching I number folder in: {day_data_path}")
        
            # Search for the matching I number folder
            for item in self.file_dao.get_directory_contents(day_data_path):
                if (self.file_dao.regex_patterns['inumber'].search(item) and
                    not re.search('reinject', item.lower())):
                    dest_path = os.path.join(day_data_path, item)
                    self.logger(f"Found matching folder: {dest_path}")
                    return dest_path
        
            self.logger(f"No matching I number folder found, returning day data path")
        else:
            self.logger(f"No I number extracted from folder name")
        
        # If no matching folder found, return the day data path
        return os.path.dirname(base_path)

    def _get_destination_for_order_by_inum(self, i_num, current_folder=None):
        """
        Get the parent folder path for an order based on I number
        
        Raises:
            PermissionError: If unable to write to the intended directory
            ValueError: If no suitable destination can be found
        """
        self.logger(f"Finding destination for I number: {i_num}")
    
        # User-selected folder is stored in the original path
        # Get the current working folder from the processor's context
        user_selected_folder = getattr(self, 'current_data_folder', None)
        
        # If we have a user-selected folder, use it first
        if user_selected_folder and os.path.exists(user_selected_folder) and os.access(user_selected_folder, os.W_OK):
            self.logger(f"Using user-selected folder: {user_selected_folder}")
            return user_selected_folder
        
        # Preferred path pattern: P:\Data\MM.DD.YY\
        today = datetime.now().strftime('%m.%d.%y')
        preferred_path = os.path.join('P:', 'Data', today)
        
        # Try preferred path first
        if os.path.exists(preferred_path) and os.access(preferred_path, os.W_OK):
            # Search for matching I number folder within preferred path
            for item in self.file_dao.get_directory_contents(preferred_path):
                if (self.file_dao.regex_patterns['bioi_folder'].search(item) and
                    not re.search('reinject', item.lower())):
                    dest_path = os.path.join(preferred_path, item)
                    self.logger(f"Found matching folder in preferred path: {dest_path}")
                    return dest_path
            
            self.logger(f"No matching folder found, using preferred path: {preferred_path}")
            return preferred_path
        
        # If preferred path doesn't work, try other methods
        search_paths = [
            os.path.join('P:', 'Data', 'Individuals'),
            os.path.join('P:', 'Data'),
            os.path.dirname(os.getcwd())
        ]
        
        for search_path in search_paths:
            if os.path.exists(search_path) and os.access(search_path, os.W_OK):
                # Search for matching I number folder
                for item in self.file_dao.get_directory_contents(search_path):
                    if (self.file_dao.regex_patterns['inumber'].search(item) and
                        not re.search('reinject', item.lower())):
                        dest_path = os.path.join(search_path, item)
                        self.logger(f"Found matching folder in alternative path: {dest_path}")
                        return dest_path
                
                # If no matching folder, use this path
                self.logger(f"Using alternative path: {search_path}")
                return search_path
        
        # If no suitable path found, raise an error
        error_msg = f"Unable to find a writable destination for I number {i_num}"
        self.logger(error_msg)
        raise ValueError(error_msg)

    def _move_file_to_destination(self, file_path, destination_folder, normalized_name):
        """Handle file placement including reinject logic"""
        file_name = os.path.basename(file_path)
        
        # Clean filename for destination (remove braces)
        clean_brace_file_name = re.sub(r'{.*?}', '', file_name)
        target_file_path = os.path.join(destination_folder, clean_brace_file_name)
        
        # Check if file is a reinject
        is_reinject = False
        if hasattr(self, 'reinject_list') and self.reinject_list:
            is_reinject = normalized_name in [self.file_dao.normalize_filename(r) for r in self.reinject_list]
        self.logger(f"Is reinject: {is_reinject}")
        
        # Handle file placement
        if os.path.exists(target_file_path):
            # File already exists, put in alternate injections
            alt_inj_folder = os.path.join(destination_folder, "Alternate Injections")
            if not os.path.exists(alt_inj_folder):
                os.makedirs(alt_inj_folder)
                
            alt_file_path = os.path.join(alt_inj_folder, file_name)
            self.file_dao.move_file(file_path, alt_file_path)
            self.logger(f"File already exists, moved to alternate injections")
            return True
        
        elif is_reinject:
            # Handle reinjections
            raw_name_idx = self.reinject_list.index(normalized_name)
            raw_name = self.raw_reinject_list[raw_name_idx]
            
            # Check for preemptive reinject
            if '{!P}' in raw_name:
                # Preemptive reinject goes to main folder
                self.file_dao.move_file(file_path, target_file_path)
                self.logger(f"Preemptive reinject moved to main folder")
            else:
                # Regular reinject goes to alternate injections
                alt_inj_folder = os.path.join(destination_folder, "Alternate Injections")
                if not os.path.exists(alt_inj_folder):
                    os.makedirs(alt_inj_folder)
                    
                alt_file_path = os.path.join(alt_inj_folder, file_name)
                self.file_dao.move_file(file_path, alt_file_path)
                self.logger(f"Reinject moved to alternate injections")
            return True
        
        else:
            # Regular file, put in main folder
            self.file_dao.move_file(file_path, target_file_path)
            self.logger(f"File moved to main folder")
            return True

    def sort_ind_folder(self, folder_path, reinject_list, order_key, recent_inumbers):
        """Sort all files in a BioI folder using batch processing"""
        self.logger(f"Processing folder: {folder_path}")
        
        # Store reinject lists for use in methods
        self.reinject_list = reinject_list
        self.raw_reinject_list = getattr(self, 'raw_reinject_list', reinject_list)
        
        # Extract I number from the folder
        i_num = self.file_dao.get_inumber_from_name(folder_path)

        # If no I-number found, try to find it from the ab1 files
        if not i_num:
            ab1_files = self.file_dao.get_files_by_extension(folder_path, ".ab1")
            for file_path in ab1_files:
                parent_dir = os.path.basename(os.path.dirname(file_path))
                i_num = self.file_dao.get_inumber_from_name(parent_dir)
                if i_num:
                    self.logger(f"Found I number {i_num} from parent directory of AB1 file")
                    break

        # Create or find the target BioI folder first
        if i_num:
            new_folder_name = f"BioI-{i_num}"
            new_folder_path = os.path.join(os.path.dirname(folder_path), new_folder_name)
            
            # Create the new folder if it doesn't exist
            if not os.path.exists(new_folder_path):
                os.makedirs(new_folder_path)
                self.logger(f"Created new BioI folder: {new_folder_path}")
        else:
            # If no I number found, use the original folder
            new_folder_path = folder_path
            self.logger(f"No I number found, using original folder: {folder_path}")
        
        # Get all .ab1 files in the folder
        ab1_files = self.file_dao.get_files_by_extension(folder_path, ".ab1")
        self.logger(f"Found {len(ab1_files)} .ab1 files in folder")
        
        # Group files by type for batch processing
        pcr_files = {}
        control_files = []
        blank_files = []
        customer_files = []
        
        # Classify files first - FIXED VERSION with no duplicate check
        for file_path in ab1_files:
            file_name = os.path.basename(file_path)
            
            # Check for PCR files
            pcr_number = self.file_dao.get_pcr_number(file_name)
            if pcr_number:
                if pcr_number not in pcr_files:
                    pcr_files[pcr_number] = []
                pcr_files[pcr_number].append(file_path)
                continue

            # Check for blank files (check this before control files)
            if self.file_dao.is_blank_file(file_name):
                self.logger(f"Identified blank file: {file_name}")
                blank_files.append(file_path)
                continue

            # Check for control files
            if self.file_dao.is_control_file(file_name, self.config.CONTROLS):
                self.logger(f"Identified control file: {file_name}")
                control_files.append(file_path)
                continue
                
            # Must be a customer file
            customer_files.append(file_path)

        # Detailed logging for debugging
        self.logger(f"Classified {len(pcr_files)} PCR numbers, {len(control_files)} controls, {len(blank_files)} blanks, {len(customer_files)} customer files")
        
        # Log all blank files for verification
        if blank_files:
            self.logger("Blank files identified:")
            for file_path in blank_files:
                self.logger(f"  - {os.path.basename(file_path)}")
        else:
            self.logger("No blank files were identified in this folder")

        # Process each group
        # Process PCR files by PCR number
        for pcr_number, files in pcr_files.items():
            self.logger(f"Processing {len(files)} files for PCR number {pcr_number}")
            for file_path in files:
                self._sort_pcr_file(file_path, pcr_number)
        
        # Process controls - Now placing in the new BioI folder
        if control_files:
            self.logger(f"Processing {len(control_files)} control files")
            controls_folder = os.path.join(new_folder_path, "Controls")
            if not os.path.exists(controls_folder):
                os.makedirs(controls_folder)
                
            for file_path in control_files:
                target_path = os.path.join(controls_folder, os.path.basename(file_path))
                moved = self.file_dao.move_file(file_path, target_path)
                if moved:
                    self.logger(f"Moved control file {os.path.basename(file_path)} to {controls_folder}")
                else:
                    self.logger(f"Failed to move control file {os.path.basename(file_path)}")
        
        # Process blanks - Now placing in the new BioI folder
        if blank_files:
            self.logger(f"Processing {len(blank_files)} blank files")
            blank_folder = os.path.join(new_folder_path, "Blank")
            if not os.path.exists(blank_folder):
                os.makedirs(blank_folder)
                
            for file_path in blank_files:
                target_path = os.path.join(blank_folder, os.path.basename(file_path))
                moved = self.file_dao.move_file(file_path, target_path)
                if moved:
                    self.logger(f"Moved blank file {os.path.basename(file_path)} to {blank_folder}")
                else:
                    self.logger(f"Failed to move blank file {os.path.basename(file_path)}")
        
        # Process customer files (with optimized order key lookup)
        if customer_files:
            self.logger(f"Processing {len(customer_files)} customer files")
            # Build order key index first
            if self.order_key_index is None:
                self.build_order_key_index(order_key)
                
            for file_path in customer_files:
                self.sort_customer_file(file_path, order_key, recent_inumbers)
        
        # Enhanced cleanup: Check if the original folder is empty or can be safely deleted
        try:
            self._cleanup_original_folder(folder_path, new_folder_path)
        except Exception as e:
            self.logger(f"Error during folder cleanup: {e}")
        
        return new_folder_path
        
    def _cleanup_original_folder(self, original_folder, new_folder):
        """
        Enhanced cleanup to remove the original folder if all files have been processed
        """
        if original_folder == new_folder:
            self.logger("Original folder is the same as new folder, no cleanup needed")
            return
        
        # Force refresh directory cache first
        self.file_dao.get_directory_contents(original_folder, refresh=True)
        
        # Check if any .ab1 files remain in the original folder
        ab1_files = self.file_dao.get_files_by_extension(original_folder, ".ab1")
        if ab1_files:
            self.logger(f"Found {len(ab1_files)} remaining .ab1 files in original folder, cannot delete")
            
            # Log the names of remaining files for debugging
            for file_path in ab1_files:
                self.logger(f"Remaining file: {os.path.basename(file_path)}")
            
            return
        
        # Refresh directory contents
        remaining_items = self.file_dao.get_directory_contents(original_folder, refresh=True)
        
        # If completely empty, delete the folder
        if not remaining_items:
            try:
                os.rmdir(original_folder)
                self.logger(f"Deleted empty original folder: {original_folder}")
                return
            except Exception as e:
                self.logger(f"Failed to delete original folder: {e}")
                return
        
        # Try to move any remaining Control/Blank folders to the new location
        moved_all = True
        for item in list(remaining_items):  # Create a copy of the list to avoid iteration issues
            item_path = os.path.join(original_folder, item)
            
            if item in ["Controls", "Blank", "Alternate Injections"]:
                # Try to move this folder to the new location if it's not empty
                if os.path.isdir(item_path) and os.listdir(item_path):
                    try:
                        # Create target folder in new location if needed
                        target_path = os.path.join(new_folder, item)
                        if not os.path.exists(target_path):
                            os.makedirs(target_path)
                        
                        # Move all files from old to new location
                        for subitem in os.listdir(item_path):
                            old_file = os.path.join(item_path, subitem)
                            new_file = os.path.join(target_path, subitem)
                            if os.path.isfile(old_file):
                                self.file_dao.move_file(old_file, new_file)
                                self.logger(f"Moved remaining file {subitem} to {target_path}")
                        
                        # Check if folder is now empty and can be deleted
                        if not os.listdir(item_path):
                            os.rmdir(item_path)
                            self.logger(f"Deleted now-empty folder: {item}")
                    except Exception as e:
                        self.logger(f"Failed to move remaining files from {item}: {e}")
                        moved_all = False
                        continue
                elif os.path.isdir(item_path) and not os.listdir(item_path):
                    # Empty folder - delete it
                    try:
                        os.rmdir(item_path)
                        self.logger(f"Deleted empty folder: {item}")
                    except Exception as e:
                        self.logger(f"Failed to delete empty folder {item}: {e}")
                        moved_all = False
            else:
                # Non-standard item
                self.logger(f"Found non-standard item in folder: {item}")
                moved_all = False
        
        # Refresh contents one more time after all operations
        remaining_items = self.file_dao.get_directory_contents(original_folder, refresh=True)
        
        # If we've successfully cleaned everything up, try to delete the original folder
        if not remaining_items:
            try:
                os.rmdir(original_folder)
                self.logger(f"Deleted original folder after cleaning: {original_folder}")
            except Exception as e:
                self.logger(f"Failed to delete original folder: {e}")
        else:
            self.logger(f"Unable to clean up original folder. {len(remaining_items)} items remain.")
    
    def _sort_pcr_file(self, file_path, pcr_number):
        """Sort a PCR file to the appropriate folder"""
        file_name = os.path.basename(file_path)
        self.logger(f"Processing PCR file: {file_name} with PCR Number: {pcr_number}")
        
        # Get the day data folder
        day_data_path = os.path.dirname(os.path.dirname(file_path))
        
        # Create PCR folder name and path
        pcr_folder_name = f"FB-{pcr_number}"
        pcr_folder_path = os.path.join(day_data_path, pcr_folder_name)
        
        # Create PCR folder if it doesn't exist
        if not os.path.exists(pcr_folder_path):
            os.makedirs(pcr_folder_path)
        
        # Use the same move file logic
        normalized_name = self.file_dao.normalize_filename(file_name)
        return self._move_file_to_destination(file_path, pcr_folder_path, normalized_name)

    def _sort_control_file(self, file_path):
        """Sort a control file to the Controls folder"""
        parent_folder = os.path.dirname(file_path)
        controls_folder = os.path.join(parent_folder, "Controls")
        
        # Create Controls folder if it doesn't exist
        if not os.path.exists(controls_folder):
            os.makedirs(controls_folder)
        
        # Just move the file directly (no special handling needed)
        target_path = os.path.join(controls_folder, os.path.basename(file_path))
        return self.file_dao.move_file(file_path, target_path)

    def _sort_blank_file(self, file_path):
        """Sort a blank file to the Blank folder"""
        # Get parent BioI folder first, not just the immediate parent folder
        immediate_parent = os.path.dirname(file_path)
        file_name = os.path.basename(file_path)
        
        # Get or create the BioI folder
        i_num = self.file_dao.get_inumber_from_name(immediate_parent)
        if i_num:
            bioi_folder_name = f"BioI-{i_num}"
            bioi_folder_path = os.path.join(os.path.dirname(immediate_parent), bioi_folder_name)
            
            # Ensure BioI folder exists
            if not os.path.exists(bioi_folder_path):
                os.makedirs(bioi_folder_path)
            
            # Create Blank folder inside BioI folder
            blank_folder = os.path.join(bioi_folder_path, "Blank")
            if not os.path.exists(blank_folder):
                os.makedirs(blank_folder)
            
            # Move file to Blank folder
            target_path = os.path.join(blank_folder, file_name)
            return self.file_dao.move_file(file_path, target_path)
        else:
            # Fallback to original logic if no I number found
            parent_folder = immediate_parent
            blank_folder = os.path.join(parent_folder, "Blank")
            if not os.path.exists(blank_folder):
                os.makedirs(blank_folder)
            target_path = os.path.join(blank_folder, file_name)
            return self.file_dao.move_file(file_path, target_path)

    def _rename_processed_folder(self, folder_path):
        """Rename the folder after processing"""
        i_num = self.file_dao.get_inumber_from_name(folder_path)
        if i_num:
            new_folder_name = f"BioI-{i_num}"
            new_folder_path = os.path.join(os.path.dirname(folder_path), new_folder_name)
            
            # Only rename if the new name doesn't already exist
            if not os.path.exists(new_folder_path) and os.path.basename(folder_path) != new_folder_name:
                try:
                    os.rename(folder_path, new_folder_path)
                    self.logger(f"Renamed folder to: {new_folder_name}")
                    return True
                except Exception as e:
                    self.logger(f"Failed to rename folder: {e}")
        
        return False
    
    def get_todays_inumbers_from_folder(self, path):
        """Get I numbers and folder paths from the selected folder"""
        i_numbers = []
        bioi_folders = []
        
        for item in self.file_dao.get_directory_contents(path):
            item_path = os.path.join(path, item)
            if os.path.isdir(item_path):
                # Check if it's a BioI folder
                if self.file_dao.regex_patterns['inumber'].search(item):
                    i_num = self.file_dao.get_inumber_from_name(item)
                    if i_num and i_num not in i_numbers:
                        i_numbers.append(i_num)
                
                # Get BioI folders before sorting (avoid reinject folders)
                if (self.file_dao.regex_patterns['bioi_folder'].search(item) and
                    'reinject' not in item.lower()):
                    bioi_folders.append(item_path)
        
        return i_numbers, bioi_folders

    def get_reinject_list(self, i_numbers, reinject_path=None):
        """Get list of reactions that are reinjects - optimized version"""
        import re
        import os
        import numpy as np
        
        reinject_list = []
        raw_reinject_list = []
        
        # Cache of processed text files to avoid reprocessing
        processed_files = set()
        
        # Process text files only once
        spreadsheets_path = r'G:\Lab\Spreadsheets'
        abi_path = os.path.join(spreadsheets_path, 'Individual Uploaded to ABI')
        
        # Build a list of potential reinject files first
        reinject_files = []
        
        # Scan directories once
        if os.path.exists(spreadsheets_path):
            for file in self.file_dao.get_directory_contents(spreadsheets_path):
                if 'reinject' in file.lower() and file.endswith('.txt'):
                    # Check if any of our I numbers are in the filename
                    if any(i_num in file for i_num in i_numbers):
                        reinject_files.append(os.path.join(spreadsheets_path, file))
        
        if os.path.exists(abi_path):
            for file in self.file_dao.get_directory_contents(abi_path):
                if 'reinject' in file.lower() and file.endswith('.txt'):
                    # Check if any of our I numbers are in the filename
                    if any(i_num in file for i_num in i_numbers):
                        reinject_files.append(os.path.join(abi_path, file))
        
        # Process each found reinject file
        for file_path in reinject_files:
            try:
                if file_path in processed_files:
                    continue
                    
                processed_files.add(file_path)
                data = np.loadtxt(file_path, dtype=str, delimiter='\t')
                
                # Parse rows 5-101 (B6:B101 in Excel terms)
                for j in range(5, min(101, data.shape[0])):
                    if j < data.shape[0] and data.shape[1] > 1:
                        raw_reinject_list.append(data[j, 1])
                        cleaned_name = self.file_dao.clean_braces_format(data[j, 1])
                        reinject_list.append(cleaned_name)
            except Exception as e:
                self.logger(f"Error processing reinject file {file_path}: {e}")
        
        # If a reinject_path is provided, also check that Excel file
        if reinject_path and os.path.exists(reinject_path):
            try:
                import pylightxl as xl
                
                # Read the Excel file
                db = xl.readxl(reinject_path)
                sheet = db.ws('Sheet1')
                
                # Get reinject entries from Excel
                reinject_prep_list = []
                for row in range(1, sheet.maxrow + 1):
                    if sheet.maxcol >= 2:  # Ensure we have at least 2 columns
                        sample = sheet.index(row, 1)  # Use index instead of address
                        primer = sheet.index(row, 2)  # Use index instead of address
                        if sample and primer:
                            reinject_prep_list.append(sample + primer)
                
                # Convert to numpy array for easier searching
                reinject_prep_array = np.array(reinject_prep_list)
                
                # Check for partial plate reinjects by comparing against reinject_prep_array
                for i_num in i_numbers:
                    # Check all txt files for this I number
                    # Fix: Change glob pattern to valid regex pattern
                    regex_pattern = f'.*{i_num}.*txt'
                    txt_files = []
                    
                    for file in self.file_dao.get_directory_contents(spreadsheets_path):
                        if re.search(regex_pattern, file, re.IGNORECASE) and not 'reinject' in file:
                            txt_files.append(os.path.join(spreadsheets_path, file))
                    
                    if os.path.exists(abi_path):
                        for file in self.file_dao.get_directory_contents(abi_path):
                            if re.search(regex_pattern, file, re.IGNORECASE) and not 'reinject' in file:
                                txt_files.append(os.path.join(abi_path, file))
                    
                    # Check each txt file for matches against reinject_prep_array
                    for file_path in txt_files:
                        try:
                            data = np.loadtxt(file_path, dtype=str, delimiter='\t')
                            for j in range(5, min(101, data.shape[0])):
                                if j < data.shape[0] and data.shape[1] > 1:
                                    # Check if this sample is in the reinject list
                                    sample_name = data[j, 1][5:] if len(data[j, 1]) > 5 else data[j, 1]
                                    indices = np.where(reinject_prep_array == sample_name)[0]
                                    
                                    if len(indices) > 0:
                                        raw_reinject_list.append(data[j, 1])
                                        cleaned_name = self.file_dao.clean_braces_format(data[j, 1])
                                        reinject_list.append(cleaned_name)
                        except Exception as e:
                            self.logger(f"Error processing txt file {file_path}: {e}")
            
            except Exception as e:
                self.logger(f"Error processing reinject Excel file: {e}")
        
        # Store the raw_reinject_list for reference in _move_file_to_destination
        self.raw_reinject_list = raw_reinject_list
        return reinject_list

if __name__ == "__main__":
    # Simple test if run directly
    from config import MseqConfig
    from file_system_dao import FileSystemDAO
    
    config = MseqConfig()
    file_dao = FileSystemDAO(config)
    processor = FolderProcessor(file_dao, None, config)
    
    # Test folder processing
    test_path = os.getcwd()
    print(f"Testing with folder: {test_path}")
    
    # Test getting I number from a folder name
    test_folder_name = "BioI-12345_Customer_67890"
    i_num = processor.file_dao.get_inumber_from_name(test_folder_name)
    print(f"I number from {test_folder_name}: {i_num}")