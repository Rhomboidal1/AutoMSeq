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
        self.logger = logger or print  # Default to print if no logger provided
    
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
    
    def check_order_status(self, folder):
        """Check if an order folder has been processed, has braces, and has ab1 files"""
        was_mseqed = False
        has_braces = False
        has_ab1_files = False
        
        # Check for mSeq artifacts
        current_artifacts = []
        for item in os.listdir(folder):
            if item in self.config.MSEQ_ARTIFACTS:
                current_artifacts.append(item)
            
            # Check for braces in ab1 files
            if item.endswith('.ab1'):
                has_ab1_files = True
                if '{' in item or '}' in item:
                    has_braces = True
        
        was_mseqed = set(current_artifacts) == self.config.MSEQ_ARTIFACTS
        return was_mseqed, has_braces, has_ab1_files
    
    def get_order_number(self, folder_name):
        """Extract order number from folder name"""
        match = re.search('_\\d+', folder_name)
        if match:
            return re.search('\\d+', match.group(0)).group(0)
        return ''
    
    def get_i_number(self, folder_name):
        """Extract I number from folder name"""
        match = re.search('bioi-\\d+', str(folder_name).lower())
        if match:
            bio_string = match.group(0)
            return re.search('\\d+', bio_string).group(0)
        return None
    
    def get_destination_for_order(self, order_folder, base_path):
        """Determine the correct destination for an order folder"""
        folder_name = os.path.basename(order_folder)
        i_num = self.get_i_number(folder_name)
        
        if i_num:
            # Navigate up one level to get to the day folder
            day_data_path = os.path.dirname(base_path)
            
            # Search for the matching I number folder
            for item in os.listdir(day_data_path):
                if (re.search(f'bioi-{i_num}', item.lower()) and 
                    not re.search('reinject', item.lower())):
                    return os.path.join(day_data_path, item)
        
        # If no matching folder found, return the day data path
        return os.path.dirname(base_path)
    
    def process_bio_folder(self, folder):
        """Process a BioI folder (specialized for IND)"""
        self.logger(f"Processing BioI folder: {os.path.basename(folder)}")
        
        # Get all order folders in this BioI folder
        order_folders = self.file_dao.get_folders(folder, r'bioi-\d+_.+_\d+')
        
        for order_folder in order_folders:
            # Skip Andreev's orders for mSeq processing
            if 'andreev' in order_folder.lower():
                continue
                
            order_number = self.get_order_number(os.path.basename(order_folder))
            ab1_files = self.file_dao.get_files_by_extension(order_folder, '.ab1')
            
            # Check order status
            was_mseqed, has_braces, has_ab1_files = self.check_order_status(order_folder)
            
            if not was_mseqed and not has_braces:
                # Process if we have the right number of ab1 files
                expected_count = self._get_expected_file_count(order_number)
                
                if len(ab1_files) == expected_count:
                    if has_ab1_files:
                        self.ui_automation.process_folder(order_folder)
                        self.logger(f"mSeq completed: {os.path.basename(order_folder)}")
                else:
                    # Move to Not Ready folder if incomplete
                    not_ready_path = os.path.join(os.path.dirname(folder), "IND Not Ready")
                    self.file_dao.create_folder_if_not_exists(not_ready_path)
                    self.file_dao.move_folder(order_folder, not_ready_path)
                    self.logger(f"Order moved to Not Ready: {os.path.basename(order_folder)}")
    
    def process_plate_folder(self, folder):
        """Process a plate folder"""
        self.logger(f"Processing plate folder: {os.path.basename(folder)}")
        
        # Check if folder contains FSA files (skip if it does)
        if self.file_dao.contains_file_type(folder, '.fsa'):
            self.logger(f"Skipping folder with FSA files: {os.path.basename(folder)}")
            return
        
        # Check if already processed
        was_mseqed = self.is_mseq_processed(folder)
        
        if not was_mseqed:
            # Process with mSeq
            self.ui_automation.process_folder(folder)
            self.logger(f"mSeq completed: {os.path.basename(folder)}")
    
    def process_wildcard_folder(self, folder):
        """Process any folder"""
        self.logger(f"Processing folder: {os.path.basename(folder)}")
        
        # Skip folders with FSA files
        if self.file_dao.contains_file_type(folder, '.fsa'):
            self.logger(f"Skipping folder with FSA files: {os.path.basename(folder)}")
            return
        
        # Check if already processed
        was_mseqed = self.is_mseq_processed(folder)
        
        if not was_mseqed:
            # Process with mSeq
            self.ui_automation.process_folder(folder)
            self.logger(f"mSeq completed: {os.path.basename(folder)}")
    
    def process_order_folder(self, order_folder, data_folder_path):
        """Process an order folder"""
        self.logger(f"Processing order folder: {os.path.basename(order_folder)}")
        
        # Skip Andreev's orders for mSeq processing
        if 'andreev' in order_folder.lower():
            order_number = self.get_order_number(os.path.basename(order_folder))
            ab1_files = self.file_dao.get_files_by_extension(order_folder, '.ab1')
            expected_count = self._get_expected_file_count(order_number)
            
            # For Andreev's orders, just check if complete to move back if needed
            if len(ab1_files) == expected_count:
                # If we're processing from IND Not Ready, move it back
                if os.path.basename(os.path.dirname(order_folder)) == "IND Not Ready":
                    destination = self.get_destination_for_order(order_folder, data_folder_path)
                    self.file_dao.move_folder(order_folder, destination)
                    self.logger(f"Andreev's order moved back: {os.path.basename(order_folder)}")
            return
            
        order_number = self.get_order_number(os.path.basename(order_folder))
        ab1_files = self.file_dao.get_files_by_extension(order_folder, '.ab1')
        
        # Check order status
        was_mseqed, has_braces, has_ab1_files = self.check_order_status(order_folder)
        
        # Process based on status
        if not was_mseqed and not has_braces:
            expected_count = self._get_expected_file_count(order_number)
            
            if len(ab1_files) == expected_count and has_ab1_files:
                self.ui_automation.process_folder(order_folder)
                self.logger(f"mSeq completed: {os.path.basename(order_folder)}")
                
                # If processing from IND Not Ready, move it back
                if os.path.basename(os.path.dirname(order_folder)) == "IND Not Ready":
                    destination = self.get_destination_for_order(order_folder, data_folder_path)
                    self.file_dao.move_folder(order_folder, destination)
            else:
                # Move to Not Ready if incomplete
                not_ready_path = os.path.join(os.path.dirname(data_folder_path), "IND Not Ready")
                self.file_dao.create_folder_if_not_exists(not_ready_path)
                self.file_dao.move_folder(order_folder, not_ready_path)
                self.logger(f"Order moved to Not Ready: {os.path.basename(order_folder)}")
        
        # If already mSeqed but in IND Not Ready, move it back
        elif was_mseqed and os.path.basename(os.path.dirname(order_folder)) == "IND Not Ready":
            destination = self.get_destination_for_order(order_folder, data_folder_path)
            self.file_dao.move_folder(order_folder, destination)
            self.logger(f"Processed order moved back: {os.path.basename(order_folder)}")
    
    def process_pcr_folder(self, folder):
        """Process a PCR folder"""
        self.logger(f"Processing PCR folder: {os.path.basename(folder)}")
        
        # Check if already processed
        was_mseqed, has_braces, has_ab1_files = self.check_order_status(folder)
        
        if not was_mseqed and not has_braces and has_ab1_files:
            self.ui_automation.process_folder(folder)
            self.logger(f"mSeq completed: {os.path.basename(folder)}")
        else:
            self.logger(f"mSeq NOT completed: {os.path.basename(folder)}")
    
    def _get_expected_file_count(self, order_number):
        """Get expected number of files for an order based on the order key"""
        # Load the order key file
        order_key = self.file_dao.load_order_key(self.config.KEY_FILE_PATH)
        if order_key is None:
            self.logger(f"Warning: Could not load order key file, unable to verify count for order {order_number}")
            return 0
        
        # Count matching entries for this order number
        count = 0
        for row in order_key:
            if str(row[0]) == str(order_number):
                count += 1
        
        return count

    def _get_order_sample_names(self, order_number):
        """Get sample names for an order"""
        order_key = self.file_dao.load_order_key(self.config.KEY_FILE_PATH)
        if order_key is None:
            return []
        
        sample_names = []
        
        # Find indices of rows matching this order number
        indices = np.where((order_key == str(order_number)))
        
        for i in range(len(indices[0])):
            # Get the sample name and adjust characters
            sample_name = self.file_dao.adjust_abi_chars(order_key[indices[0][i], 3])
            sample_names.append(sample_name)
        
        return sample_names
   
    def sort_controls(self, folder_path):
        """Sort control files into Controls folder"""
        controls_path = os.path.join(folder_path, self.config.CONTROLS_FOLDER)
        self.file_dao.create_folder_if_not_exists(controls_path)
        
        for file in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file)
            if not os.path.isfile(file_path) or not file.endswith(self.config.AB1_EXTENSION):
                continue
                
            if self.file_dao.is_control_file(file, self.config.CONTROLS):
                self.file_dao.move_file(file_path, os.path.join(controls_path, file))
                self.logger(f"Moved control file: {file}")
    
    def sort_blanks(self, folder_path):
        """Sort blank files into Blank folder"""
        blank_path = os.path.join(folder_path, self.config.BLANK_FOLDER)
        self.file_dao.create_folder_if_not_exists(blank_path)
        
        for file in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file)
            if not os.path.isfile(file_path) or not file.endswith(self.config.AB1_EXTENSION):
                continue
                
            if self.file_dao.is_blank_file(file):
                self.file_dao.move_file(file_path, os.path.join(blank_path, file))
                self.logger(f"Moved blank file: {file}")
    
    # Zip methods
    def zip_order_folder(self, folder_path, include_txt=True):
        """Zip ab1 and txt files in an order folder"""
        if self.file_dao.check_for_zip(folder_path):
            self.logger(f"Zip already exists for {os.path.basename(folder_path)}")
            return False
        
        # Check if ab1 files exist
        ab1_files = self.file_dao.get_files_by_extension(folder_path, self.config.AB1_EXTENSION)
        if not ab1_files:
            self.logger(f"No ab1 files found in {os.path.basename(folder_path)}")
            return False
        
        # Check if txt files exist (if needed)
        has_txt = True
        if include_txt:
            txt_count = 0
            for txt_ext in self.config.TEXT_FILES:
                txt_count += len(self.file_dao.get_files_by_extension(folder_path, txt_ext))
            has_txt = txt_count == 5  # All 5 text files must be present
        
        if not has_txt and include_txt:
            self.logger(f"Missing txt files in {os.path.basename(folder_path)}")
            return False
        
        # Create zip file
        zip_name = os.path.basename(folder_path) + self.config.ZIP_EXTENSION
        zip_path = os.path.join(folder_path, zip_name)
        
        extensions = [self.config.AB1_EXTENSION]
        if include_txt:
            extensions.extend(self.config.TEXT_FILES)
        
        # Create the zip
        success = self.file_dao.zip_files(folder_path, zip_path, file_extensions=extensions)
        
        if success:
            self.logger(f"Created zip file: {zip_name}")
            return zip_path
        
        return False
    
    def get_todays_inumbers_from_folder(self, path):
        """Get I numbers and folder paths from the selected folder"""
        i_numbers = []
        bioi_folders = []
        
        for item in os.listdir(path):
            item_path = os.path.join(path, item)
            if os.path.isdir(item_path):
                # Check if it's a BioI folder
                if re.search('bioi-\\d+', item.lower()):
                    i_num = self.get_i_number(item)
                    if i_num and i_num not in i_numbers:
                        i_numbers.append(i_num)
                
                # Get BioI folders before sorting (avoid reinject folders)
                if (re.search('.+bioi-\\d+.+', item.lower()) and 
                    not 'reinject' in item.lower()):
                    bioi_folders.append(item_path)
        
        return i_numbers, bioi_folders

    def get_reinject_list(self, i_numbers, reinject_path=None):
        """Get list of reactions that are reinjects"""
        import re
        import os
        
        reinject_list = []
        raw_reinject_list = []
        
        # First get reinjects from txt files
        for i_num in i_numbers:
            # Check G:\Lab\Spreadsheets for reinject files
            spreadsheets_path = r'G:\Lab\Spreadsheets'
            reinject_files = []
            
            # Try to find reinject files in main spreadsheets folder
            glob_pattern = f'*{i_num}*reinject*txt'
            for file in os.listdir(spreadsheets_path):
                if re.search(glob_pattern, file, re.IGNORECASE):
                    reinject_files.append(os.path.join(spreadsheets_path, file))
            
            # Also check the "Individual Uploaded to ABI" subfolder
            abi_path = os.path.join(spreadsheets_path, 'Individual Uploaded to ABI')
            if os.path.exists(abi_path):
                for file in os.listdir(abi_path):
                    if re.search(glob_pattern, file, re.IGNORECASE):
                        reinject_files.append(os.path.join(abi_path, file))
            
            # Process each found reinject file
            for file_path in reinject_files:
                try:
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
                        sample = sheet.address(row=row, col=1)
                        primer = sheet.address(row=row, col=2)
                        if sample and primer:
                            reinject_prep_list.append(sample + primer)
                
                # Convert to numpy array for easier searching
                reinject_prep_array = np.array(reinject_prep_list)
                
                # Check for partial plate reinjects by comparing against reinject_prep_array
                for i_num in i_numbers:
                    # Check all txt files for this I number
                    glob_pattern = f'*{i_num}*txt'
                    txt_files = []
                    
                    for file in os.listdir(spreadsheets_path):
                        if re.search(glob_pattern, file, re.IGNORECASE) and not 'reinject' in file.lower():
                            txt_files.append(os.path.join(spreadsheets_path, file))
                    
                    if os.path.exists(abi_path):
                        for file in os.listdir(abi_path):
                            if re.search(glob_pattern, file, re.IGNORECASE) and not 'reinject' in file.lower():
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
        
        return reinject_list

    def remove_braces_from_filenames(self, folder_path):
        """Remove braces from all filenames in a folder"""
        for item in os.listdir(folder_path):
            if '{' in item and '}' in item:
                file_path = os.path.join(folder_path, item)
                if os.path.isfile(file_path):
                    self.file_dao.rename_file_without_braces(file_path)
                    self.logger(f"Renamed file: {item}")

    def find_zip_file(self, folder_path):
        """Find zip file in a folder"""
        for file in os.listdir(folder_path):
            if file.endswith(self.config.ZIP_EXTENSION):
                return os.path.join(folder_path, file)
        return None

    def get_zip_mod_time(self, worksheet, order_number):
        """Get zip modification time from Excel worksheet using pylightxl"""
        # If worksheet is a pylightxl worksheet
        if hasattr(worksheet, 'index'):
            for row in range(1, worksheet.maxrow + 1):
                # Check if column B (2) has the order number
                if worksheet.index(row, 2) == str(order_number):
                    # Return timestamp from column H (8)
                    return worksheet.index(row, 8)
        else:
            # Assume openpyxl worksheet (fallback)
            for row in worksheet.iter_rows(values_only=True):
                if row[1] == str(order_number):
                    return row[7]
        
        return ""

    def sort_ind_folder(self, folder, reinject_list, order_key, recent_inums):
        """Sort files in an individual folder based on order keys and reinject status"""
        self.logger(f"Sorting files in folder: {os.path.basename(folder)}")
        
        # Create order folders first
        i_number = self.get_i_number(folder)
        if i_number:
            order_folders = self._get_order_folders_for_inumber(i_number, order_key)
            
            # Create folders for each order
            for order_name in order_folders:
                order_folder_path = os.path.join(folder, order_name)
                if not os.path.exists(order_folder_path):
                    os.makedirs(order_folder_path)
                    self.logger(f"Created order folder: {order_name}")
        
        # Sort files
        for file in os.listdir(folder):
            file_path = os.path.join(folder, file)
            if not os.path.isfile(file_path):
                continue
                
            # Only process .ab1 files
            if not file.endswith('.ab1'):
                continue
                
            # Process controls
            clean_file_name = self.file_dao.adjust_abi_chars(file)[:-4]  # Remove .ab1
            clean_file_name = self.file_dao.neutralize_suffixes(clean_file_name)
            clean_file_name = re.sub(r'{.*?}', '', clean_file_name)
            
            if clean_file_name in self.config.CONTROLS:
                # Move to Controls folder
                controls_path = os.path.join(folder, self.config.CONTROLS_FOLDER)
                if not os.path.exists(controls_path):
                    os.makedirs(controls_path)
                
                dest_path = os.path.join(controls_path, file)
                self.file_dao.move_file(file_path, dest_path)
                self.logger(f"Moved control file: {file}")
                continue
                
            # Process blanks
            if clean_file_name == '':
                # Move to Blank folder
                blank_path = os.path.join(folder, self.config.BLANK_FOLDER)
                if not os.path.exists(blank_path):
                    os.makedirs(blank_path)
                    
                dest_path = os.path.join(blank_path, file)
                self.file_dao.move_file(file_path, dest_path)
                self.logger(f"Moved blank file: {file}")
                continue
            
            # Process PCR files
            pcr_number = self.file_dao.get_pcr_number(file)
            if pcr_number:
                # Create PCR folder if needed
                pcr_folder_name = self._get_pcr_folder_name(pcr_number)
                pcr_folder_path = os.path.join(os.path.dirname(folder), pcr_folder_name)
                
                if not os.path.exists(pcr_folder_path):
                    os.makedirs(pcr_folder_path)
                
                # Check if this is a reinject
                is_reinject = clean_file_name in reinject_list
                
                # If file already exists or is reinject, put in Alternate Injections
                clean_brace_file = re.sub(r'{.*?}', '', file)
                if os.path.exists(os.path.join(pcr_folder_path, clean_brace_file)) or is_reinject:
                    alt_inj_path = os.path.join(pcr_folder_path, self.config.ALT_INJECTIONS_FOLDER)
                    if not os.path.exists(alt_inj_path):
                        os.makedirs(alt_inj_path)
                    
                    dest_path = os.path.join(alt_inj_path, file)
                    self.file_dao.move_file(file_path, dest_path)
                    self.logger(f"Moved PCR file to Alternate Injections: {file}")
                else:
                    # Move to PCR folder with clean name
                    dest_path = os.path.join(pcr_folder_path, clean_brace_file)
                    self.file_dao.move_file(file_path, dest_path)
                    self.logger(f"Moved PCR file: {file}")
                
                continue
            
            # Process customer files - find matching order
            destination = self._determine_sort_destination(
                file, clean_file_name, order_key, i_number
            )
            
            if destination:
                # Check if file already exists or is reinject
                clean_brace_file = re.sub(r'{.*?}', '', file)
                
                if os.path.exists(os.path.join(destination, clean_brace_file)) or clean_file_name in reinject_list:
                    alt_inj_path = os.path.join(destination, self.config.ALT_INJECTIONS_FOLDER)
                    if not os.path.exists(alt_inj_path):
                        os.makedirs(alt_inj_path)
                    
                    dest_path = os.path.join(alt_inj_path, file)
                    self.file_dao.move_file(file_path, dest_path)
                    self.logger(f"Moved file to Alternate Injections: {file}")
                else:
                    # Move to order folder with clean name
                    dest_path = os.path.join(destination, clean_brace_file)
                    self.file_dao.move_file(file_path, dest_path)
                    self.logger(f"Moved file: {file}")
            else:
                self.logger(f"No destination found for file: {file}")
        
        # Rename the folder to a clean format
        new_folder_name = self._update_folder_name(folder)
        if new_folder_name != folder and not os.path.exists(new_folder_name):
            os.rename(folder, new_folder_name)
            self.logger(f"Renamed folder to: {os.path.basename(new_folder_name)}")

    def _get_order_folders_for_inumber(self, i_number, order_key):
        """Get order folder names for a specific I number"""
        import numpy as np
        
        folder_list = []
        
        # Find rows matching this I number
        indices = np.where(order_key == str(i_number))[0]
        
        # Group by order number
        order_dict = {}
        for idx in indices:
            order_num = order_key[idx][2]
            acct_name = order_key[idx][1]
            
            if order_num not in order_dict:
                order_dict[order_num] = acct_name
        
        # Create folder names
        for order_num, acct_name in order_dict.items():
            folder_name = f'BioI-{i_number}_{acct_name}_{order_num}'
            if folder_name not in folder_list:
                folder_list.append(folder_name)
        
        return folder_list

    def _get_pcr_folder_name(self, pcr_number):
        """Generate PCR folder name"""
        return f'FB-PCR{pcr_number}_{pcr_number}'

    def _determine_sort_destination(self, file, clean_name, order_key, current_inum):
        """Determine where a file should be sorted to"""
        import numpy as np
        
        # Find all matching entries in order key
        indices = np.where(order_key[:, 3] == clean_name)[0]
        
        if len(indices) > 0:
            # If there's exactly one match
            if len(indices) == 1:
                idx = indices[0]
                i_num = order_key[idx][0]
                acct_name = order_key[idx][1]
                order_num = order_key[idx][2]
                
                # Create order folder name
                order_folder = f'BioI-{i_num}_{acct_name}_{order_num}'
                
                # Get parent folder path
                parent_folder = self._get_destination_for_order_by_inum(i_num)
                
                # Return full path
                return os.path.join(parent_folder, order_folder)
            
            # If there are multiple matches
            else:
                # Check if file has specific I number in braces
                specified_inum = self._get_specified_inum(file)
                
                if specified_inum:
                    # Find match with this I number
                    for idx in indices:
                        if order_key[idx][0] == specified_inum:
                            i_num = specified_inum
                            acct_name = order_key[idx][1]
                            order_num = order_key[idx][2]
                            
                            # Create order folder name
                            order_folder = f'BioI-{i_num}_{acct_name}_{order_num}'
                            
                            # Get parent folder path
                            parent_folder = self._get_destination_for_order_by_inum(i_num)
                            
                            # Return full path
                            return os.path.join(parent_folder, order_folder)
                
                # If no specific I number, try to match to current I number
                if current_inum:
                    for idx in indices:
                        if order_key[idx][0] == current_inum:
                            i_num = current_inum
                            acct_name = order_key[idx][1]
                            order_num = order_key[idx][2]
                            
                            # Create order folder name
                            order_folder = f'BioI-{i_num}_{acct_name}_{order_num}'
                            
                            # Get parent folder path
                            parent_folder = self._get_destination_for_order_by_inum(i_num)
                            
                            # Return full path
                            return os.path.join(parent_folder, order_folder)
                
                # If still no match, check if file has specific order number
                specified_order = self._get_specified_order_num(file)
                
                if specified_order:
                    for idx in indices:
                        if order_key[idx][2] == specified_order:
                            i_num = order_key[idx][0]
                            acct_name = order_key[idx][1]
                            order_num = specified_order
                            
                            # Create order folder name
                            order_folder = f'BioI-{i_num}_{acct_name}_{order_num}'
                            
                            # Get parent folder path
                            parent_folder = self._get_destination_for_order_by_inum(i_num)
                            
                            # Return full path
                            return os.path.join(parent_folder, order_folder)
        
        return None

    def _get_destination_for_order_by_inum(self, i_num):
        """Get the parent folder path for an order based on I number"""
        # Implementation will depend on your directory structure
        # This is a simplified version
        parent_dir = os.path.dirname(os.getcwd())
        return os.path.join(parent_dir, f'BioI-{i_num}')

    def _get_specified_inum(self, file_name):
        """Extract specifically tagged I number from filename"""
        match = re.search(r'{i.(\d+)}', file_name.lower())
        if match:
            return match.group(1)
        return ''

    def _get_specified_order_num(self, file_name):
        """Extract specifically tagged order number from filename"""
        match = re.search(r'{(\d+)}', file_name)
        if match:
            return match.group(1)
        return ''

    def _update_folder_name(self, folder_path):
        """Clean up folder name to standard format"""
        base_name = os.path.basename(folder_path)
        i_num = self.get_i_number(base_name)
        
        if i_num:
            new_base = f'BioI-{i_num}'
            new_path = os.path.join(os.path.dirname(folder_path), new_base)
            return new_path
        
        return folder_path

    def zip_plate_folder(self, folder_path, fsa_only=False):
        """Zip a plate folder with FSA or AB1+TXT files"""
        if self.file_dao.check_for_zip(folder_path):
            self.logger(f"Zip already exists for {os.path.basename(folder_path)}")
            return False
        
        # Set up file types to include based on fsa_only flag
        if fsa_only:
            # FSA plate - only zip FSA files
            fsa_files = self.file_dao.get_files_by_extension(folder_path, self.config.FSA_EXTENSION)
            if not fsa_files:
                self.logger(f"No FSA files found in {os.path.basename(folder_path)}")
                return False
                
            extensions = [self.config.FSA_EXTENSION]
        else:
            # Normal plate - zip AB1 and TXT files
            ab1_files = self.file_dao.get_files_by_extension(folder_path, self.config.AB1_EXTENSION)
            if not ab1_files:
                self.logger(f"No AB1 files found in {os.path.basename(folder_path)}")
                return False
                
            # Check for TXT files
            has_txt = True
            txt_count = 0
            for txt_ext in self.config.TEXT_FILES:
                txt_count += len(self.file_dao.get_files_by_extension(folder_path, txt_ext))
            
            has_txt = txt_count == 5  # Need all 5 TXT files
            
            if not has_txt:
                self.logger(f"Missing TXT files in {os.path.basename(folder_path)}")
                return False
                
            extensions = [self.config.AB1_EXTENSION] + self.config.TEXT_FILES
        
        # Create zip file
        zip_name = os.path.basename(folder_path) + self.config.ZIP_EXTENSION
        zip_path = os.path.join(folder_path, zip_name)
        
        # Create the zip
        success = self.file_dao.zip_files(folder_path, zip_path, file_extensions=extensions)
        
        if success:
            self.logger(f"Created zip file: {zip_name}")
            return zip_path
        
        return False

    def validate_zip_contents(self, zip_path, i_number, order_number, order_key):
        """Validate zip contents against order key"""
        import numpy as np
        
        # Get zip contents
        zip_contents = self.file_dao.get_zip_contents(zip_path)
        if not zip_contents:
            self.logger(f"Empty or invalid zip file: {os.path.basename(zip_path)}")
            return None
        
        # Find expected order items
        try:
            indices = np.where(order_key[:, 0] == str(order_number))[0]
            expected_items = []
            
            for idx in indices:
                # Get raw and adjusted sample names
                raw_name = order_key[idx][3]
                adjusted_name = self.file_dao.adjust_abi_chars(raw_name)
                expected_items.append({
                    'raw_name': raw_name,
                    'adjusted_name': adjusted_name
                })
            
            # Analyze contents
            matches = []
            mismatches_in_zip = []
            mismatches_in_order = []
            txt_files = []
            
            # Check all AB1 files in zip against expected items
            for zip_item in zip_contents:
                if zip_item.endswith(self.config.AB1_EXTENSION):
                    # Clean zip item name
                    clean_name = self.file_dao.neutralize_suffixes(zip_item)
                    clean_name = self.file_dao.remove_extension(clean_name, self.config.AB1_EXTENSION)
                    clean_name = re.sub(r'{.*?}', '', clean_name)
                    
                    # Look for match in expected items
                    found = False
                    for expected in expected_items:
                        if clean_name == expected['adjusted_name']:
                            matches.append({
                                'raw_name': expected['raw_name'],
                                'file_name': zip_item
                            })
                            expected['matched'] = True
                            found = True
                            break
                    
                    if not found:
                        mismatches_in_zip.append(zip_item)
                
                # Track TXT files
                for txt_ext in self.config.TEXT_FILES:
                    if zip_item.endswith(txt_ext):
                        txt_files.append(txt_ext)
            
            # Find expected items without matches
            for expected in expected_items:
                if not expected.get('matched', False):
                    mismatches_in_order.append({
                        'raw_name': expected['raw_name']
                    })
            
            # Return validation results
            return {
                'matches': matches,
                'mismatches_in_zip': mismatches_in_zip,
                'mismatches_in_order': mismatches_in_order,
                'txt_files': txt_files,
                'match_count': len(matches),
                'mismatch_count': len(mismatches_in_zip) + len(mismatches_in_order),
                'txt_count': len(txt_files),
                'expected_count': len(expected_items)
            }
        
        except Exception as e:
            self.logger(f"Error validating zip contents: {e}")
            return None
    

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
    i_num = processor.get_i_number(test_folder_name)
    print(f"I number from {test_folder_name}: {i_num}")
    
    # Test getting order number from a folder name
    order_num = processor.get_order_number(test_folder_name)
    print(f"Order number from {test_folder_name}: {order_num}")