import os
import sys
import datetime
import traceback
import subprocess
from os import listdir, path as OsPath, mkdir
from shutil import move as ShutilMove
from re import search, sub
from time import sleep as TimeSleep
from datetime import date
from numpy import loadtxt, where
from subprocess import CalledProcessError, run as SubprocessRun
from tkinter import filedialog

# Setup for 32-bit Python requirement
import warnings
sys.coinit_flags = 2
warnings.filterwarnings("ignore", message="Apply externally defined coinit_flags*", category=UserWarning)
warnings.filterwarnings("ignore", message="Revert to STA COM threading mode", category=UserWarning)

# Check if running in 64-bit Python
if sys.maxsize > 2**32:
    # Path to 32-bit Python
    py32_path = r"C:\Python312-32\python.exe"
    
    if os.path.exists(py32_path):
        # Get the full path of the current script
        script_path = os.path.abspath(__file__)
        
        # Re-run this script with 32-bit Python and exit current process
        subprocess.run([py32_path, script_path])
        sys.exit(0)
    else:
        print("32-bit Python not found at", py32_path)
        print("Continuing with 64-bit Python (may cause issues)")

# Global variable to track first browse operation
firstTimeBrowsingMseq = True

#############################################
# LOGGING FUNCTIONS
#############################################

def write_error_log(log_messages=None):
    """
    Creates a log file with timestamp and any error messages.
    
    Args:
        log_messages: List of strings or a single string message to log
    """
    try:
        # Create logs directory if it doesn't exist
        log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
            
        # Create log filename with current date/time
        current_time = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        log_path = os.path.join(log_dir, f"mseq_log_{current_time}.txt")
        
        # Write to log file
        with open(log_path, "w") as log_file:
            log_file.write(f"=== Mseq Automation Log: {current_time} ===\n\n")
            
            # Write system info
            import platform
            log_file.write(f"OS: {platform.platform()}\n")
            log_file.write(f"Python: {platform.python_version()}\n\n")
            
            # Write any messages
            if log_messages:
                if isinstance(log_messages, list):
                    for msg in log_messages:
                        log_file.write(f"{msg}\n")
                else:
                    log_file.write(f"{log_messages}\n")
            
            # If no messages, note successful completion
            else:
                log_file.write("Program completed successfully with no errors.\n")
                
        print(f"Log file created at: {log_path}")
        return log_path
        
    except Exception as e:
        print(f"Error creating log file: {str(e)}")
        return None

def init_log_file():
    """Creates a log file at the start of program execution and returns the file path and handle"""
    try:
        # Create logs directory if it doesn't exist
        log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
            
        # Create log filename with current date/time
        current_time = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        log_path = os.path.join(log_dir, f"mseq_log_{current_time}.txt")
        
        # Open log file for writing
        log_file = open(log_path, "w")
        log_file.write(f"=== Mseq Automation Log: {current_time} ===\n\n")
        
        # Write system info
        import platform
        log_file.write(f"OS: {platform.platform()}\n")
        log_file.write(f"Python: {platform.python_version()}\n\n")
        log_file.flush()
        
        print(f"Log file created at: {log_path}")
        return log_path, log_file
        
    except Exception as e:
        print(f"Error creating log file: {str(e)}")
        return None, None

def log_message(log_file, message):
    """Logs a message to the specified file handle"""
    if log_file and not log_file.closed:
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        log_file.write(f"[{timestamp}] {message}\n")
        log_file.flush()  # Force write to disk
        print(message)

def enhanced_logging(message, log_file=None):
    """Helper function to log messages to both console and log file"""
    timestamp = datetime.datetime.now().strftime("%H:%M:%S")
    formatted_message = f"[{timestamp}] NAV: {message}"
    print(formatted_message)
    
    # Log to file if provided
    if log_file and not log_file.closed:
        log_file.write(formatted_message + "\n")
        log_file.flush()  # Force write to disk

#############################################
# UTILITY FUNCTIONS
#############################################

def GetFormattedDate():
    """Returns today's date in format m-d-yyyy"""
    today = date.today()
    #gets rid of leading zeros
    if today.strftime('%m')[:1]=='0':
        m = today.strftime('%m')[1:]
    else:
        m = today.strftime('%m')

    if today.strftime('%d')[:1]=='0':
        d = today.strftime('%d')[1:]
    else:
        d = today.strftime('%d')
    y = today.strftime('%Y')

    formatteDate= f'{m}-{d}-{y}'
    return formatteDate

def CleanBracesFormat(aList):
    """Returns a version of the string without anything contained in {} 
    ex: {01A}Sample_Primer{5_25}{I-12345} returns Sample_Primer"""
    i = 0
    newList = []
    while i < len(aList):
        newList.append(sub('{.*?}', '', NeutralizeSuffixes(aList[i])))
        i+=1
    return newList

def get_windows_version():
    """Detect whether the system is running Windows 10 or Windows 11"""
    # TEMPORARY OVERRIDE FOR TESTING
    return "Windows 11"  # Force Windows 11 mode for testing
    
    # Original detection code (will not be reached due to return above)
    """Detect whether the system is running Windows 10 or Windows 11"""
    import platform
    import sys
    
    # Get Windows version info
    windows_version = sys.getwindowsversion()
    build_number = windows_version.build
    
    # Windows 11 is Windows 10 version 21H2 build 22000 or higher
    if build_number >= 22000:
        return "Windows 11"
    else:
        return "Windows 10"

def get_virtual_folder_name():
    """Gets the name of the virtual folder ('This PC' in English versions)"""
    import win32com.client
    shell = win32com.client.Dispatch("Shell.Application")
    namespace = shell.Namespace(0x11)  # CSIDL_DRIVES
    folder_name = namespace.Title
    return folder_name

#############################################
# I-NUMBER AND FOLDER FUNCTIONS
#############################################

def GetINumberFromFolderName(folderName):
    """Function gets the number from the folder name: '01-01-1999_BioI-12345_1' returns '12345'"""
    if search('bioi-\\d+', str(folderName).lower()):
        BioString = search('bioi-\\d+', str(folderName).lower()).group(0)
        Inumber = search('\\d+', BioString).group(0)
        return Inumber
    else:
        return None

def GetOrderNumberFromFolderName(folder):
    """Function finds and returns just the order number of the BioI folder name"""
    orderNum = ''
    if search('.+_\\d+', folder):
        orderNum = search('_\\d+', folder).group(0)
        orderNum = search('\\d+', orderNum).group(0)
        return orderNum
    else:
        return ''

def GetNumbers(files):
    numbers = []
    for file in files:
        numbers.append(GetINumberFromFolderName(file))
    return numbers

def GetInumberFolders(path):
    INumberPaths = []
    for item in listdir(path):                              
        if OsPath.isdir(f'{path}\\{item}'):                    
            if search('bioi-\\d+', item.lower()) and (item.lower()[:4]=='bioi') and (search('reinject', item.lower())==None) and (search('bioi-\\d+_.+_\\d+', item.lower())==None):       
                INumberPaths.append(f'{path}\\{item}')
    return  INumberPaths

def GetImmediateOrders(path):
    """Returns a list of paths to order folders in the given path"""
    orderPaths = []
    for item in listdir(path):
        if OsPath.isdir(f'{path}\\{item}'):
            if search('bioi-\\d+_.+_\\d+', item.lower()) and not 'reinject' in item.lower():
                orderPaths.append(f'{path}\\{item}')
    return orderPaths

def GetPCRFolders(path):
    """Function gets PCR folders given a directory"""
    pcrpaths = []
    for item in listdir(path):
        if OsPath.isdir(f'{path}\\{item}'):
            if search('fb-pcr\\d+_\\d+', item.lower()):
                pcrpaths.append(f'{path}\\{item}')
    return pcrpaths

def GetOrderFolders(plateFolder):
    """Function gets full paths of the order folders given a BioI folder. 
    Returns list of full paths of order folders."""
    orderFolderPaths = []
    for item in listdir(plateFolder):
        if OsPath.isdir(f'{plateFolder}\\{item}'):
            if search('bioi-\\d+_.+_\\d+', item.lower()) and (search('reinject', item.lower())==None):            
                orderFolderPaths.append(f'{plateFolder}\\{item}')   
    return orderFolderPaths

def GetDestination(order, path):
    destPath = ''

    #order folders look like: BioI-12345_Name_123456
    if search('bioi-\\d+', order.lower()):
        iNum = search('bioi-\\d+', order.lower()).group(0)
        iNum = search('\\d+', iNum).group(0)

        dataFolderPath = sub(r'/','\\\\', path)
        dayData = dataFolderPath.split('\\')
        dayDataPath = '\\'.join(dayData[:len(dayData)-1])

        for item in listdir(dayDataPath):
            if search(f'bioi-{iNum}', item.lower()) and (search('reinject', item.lower())==None):
                destPath = f'{dayDataPath}\\{item}\\'
                break

    if destPath == '':
        destPath = dayDataPath + '\\'
    
    return destPath

#############################################
# FILE AND ORDER FUNCTIONS
#############################################

def AdjustABIChars(fileName):
    """Function changes characters in file name just as they are changed when making txt files from excel spreadsheets."""
    newFileName = fileName
    newFileName = newFileName.replace(' ', '')
    newFileName = newFileName.replace('+', '&')
    newFileName = newFileName.replace("*", "-")
    newFileName = newFileName.replace("|", "-")
    newFileName = newFileName.replace("/", "-")
    newFileName = newFileName.replace("\\", "-")
    newFileName = newFileName.replace(":", "-")
    newFileName = newFileName.replace("\"", "")
    newFileName = newFileName.replace("\'", "")
    newFileName = newFileName.replace("<", "-")
    newFileName = newFileName.replace(">", "-")
    newFileName = newFileName.replace("?", "")
    newFileName = newFileName.replace(",", "")
    return newFileName

def NeutralizeSuffixes(fileName):
    newFileName = fileName
    newFileName = newFileName.replace('_Premixed', '')
    newFileName = newFileName.replace('_RTI', '')
    return newFileName

def AdjustFullKeyToABIChars():
    i = 0
    while i < len(key):
        key[i][3] = AdjustABIChars(key[i][3])
        i+=1

def GetOrderList(number):
    orderSampleNames = []
    orderIndexes = where((key == str(number)))
    i = 0 
    while i < len(orderIndexes[0]):
        adjustedName = AdjustABIChars(key[orderIndexes[0][i], 3])
        orderSampleNames.append(adjustedName)
        i+=1
    return orderSampleNames
        
def GetAllab1Files(folder):
    ab1FilePaths = []
    for item in listdir(folder):
        if OsPath.isfile(f'{folder}\\{item}'):
            if item.endswith('.ab1'):
                ab1FilePaths.append(f'{folder}\\{item}')
    return ab1FilePaths

def OrderInReinjects(order):
    inReinjectList = False
    for rxn in order:
        if rxn in completeReinjectList:
            inReinjectList = True
            break
    return inReinjectList

def CheckOrder(orderfolder):
    """Function checks the order folder to tell if it was mseqed yet, if any ab1 files have braces, 
    and ensures there are ab1 files present."""
    wasMSeqed = False
    areBraces = False
    ab1s = False
    currentProj = []
    mseqSet = {'chromat_dir', 'edit_dir', 'phd_dir', 'mseq4.ini'}
    for item in listdir(orderfolder):
        if item in mseqSet:
            currentProj.append(item)
        if ('{' in item or '}' in item) and item.endswith('.ab1'): 
            areBraces = True
        if item.endswith('.ab1'):
            ab1s = True
    projSet = set(currentProj)

    if mseqSet == projSet: 
        wasMSeqed = True

    return wasMSeqed, areBraces, ab1s

def CheckFor5txt(orderPath):
    fivetxts = 0
    all5 = False
    for item in listdir(orderPath):
        if OsPath.isfile(f'{orderPath}\\{item}'):
            if item.endswith('.raw.qual.txt'):
                fivetxts +=1
            elif item.endswith('.raw.seq.txt'):
                fivetxts +=1
            elif item.endswith('.seq.info.txt'):
                fivetxts +=1
            elif item.endswith('.seq.qual.txt'):
                fivetxts +=1
            elif item.endswith('.seq.txt'):
                fivetxts +=1
    all5 = True if (fivetxts == 5) else False
    return all5

#############################################
# MSEQ APPLICATION FUNCTIONS
#############################################

def MseqOrder(orderpath, log_file=None):
    """
    Enhanced MSeq order processing with better error handling and logging
    """
    global firstTimeBrowsingMseq
    from pywinauto.keyboard import send_keys
    from pywinauto import timings
    from time import sleep as TimeSleep
    
    try:
        if log_file:
            enhanced_logging(f"Starting mSeq for {orderpath}", log_file)
        
        # Get mSeq application
        app, mseqMainWindow = GetMseq()
        mseqMainWindow.set_focus()
        
        # Create new project with Ctrl+N
        if log_file:
            enhanced_logging("Sending Ctrl+N to create new project", log_file)
        send_keys('^n')
        
        # Wait for Browse dialog to appear
        browse_success = False
        for attempt in range(3):
            try:
                if log_file:
                    enhanced_logging(f"Waiting for Browse dialog (attempt {attempt+1})", log_file)
                if timings.wait_until(timeout=10, retry_interval=0.5, func=lambda: IsBrowseDialogOpen(app), value=True):
                    browse_success = True
                    if log_file:
                        enhanced_logging("Browse dialog appeared", log_file)
                    break
            except Exception as e:
                if log_file:
                    enhanced_logging(f"Browse dialog wait error: {str(e)}", log_file)
                TimeSleep(1.0)
                if attempt < 2:  # Don't retry on last attempt
                    mseqMainWindow.set_focus()
                    send_keys('^n')
        
        if not browse_success:
            if log_file:
                enhanced_logging("Browse dialog never appeared - aborting", log_file)
            return False
        
        # Get browse dialog window
        dialogWindow = app.window(title='Browse For Folder')
        
        # Handle timing for first browse operation
        if firstTimeBrowsingMseq:
            firstTimeBrowsingMseq = False
            TimeSleep(1.5)  # Extra time for first browse
            if log_file:
                enhanced_logging("First-time browsing - using extended wait", log_file)
        else:
            TimeSleep(0.8)
            if log_file:
                enhanced_logging("Subsequent browsing - using standard wait", log_file)
        
        # Navigate to the target folder - this uses OS-specific method
        nav_success = NavigateToFolder(dialogWindow, orderpath, log_file)
        
        if not nav_success:
            if log_file:
                enhanced_logging(f"Failed to navigate to {orderpath}", log_file)
            # Try to cancel dialog
            try:
                cancel_button = dialogWindow.child_window(title="Cancel", class_name="Button")
                if cancel_button.exists():
                    cancel_button.click_input()
                    if log_file:
                        enhanced_logging("Clicked Cancel button", log_file)
                else:
                    send_keys('{ESC}')
                    if log_file:
                        enhanced_logging("Sent ESC key", log_file)
            except:
                send_keys('{ESC}')
            return False
        
        # Click OK button
        if log_file:
            enhanced_logging("Navigation successful, clicking OK button", log_file)
            
        try:
            okDiaglogButton = app.BrowseForFolder.child_window(title="OK", class_name="Button")
            okDiaglogButton.click_input()
            if log_file:
                enhanced_logging("Clicked OK button", log_file)
        except Exception as ok_error:
            if log_file:
                enhanced_logging(f"Error clicking OK: {str(ok_error)}", log_file)
            # Try alternative OK button approach
            try:
                ok_button = dialogWindow.child_window(title="OK", class_name="Button")
                ok_button.click_input()
                if log_file:
                    enhanced_logging("Clicked OK button (alternative method)", log_file)
            except:
                return False
        
        # Wait for and handle Preferences window
        if log_file:
            enhanced_logging("Waiting for Preferences window", log_file)
            
        try:
            if timings.wait_until(timeout=10, retry_interval=0.5, func=lambda: IsPreferencesOpen(app), value=True):
                mseqPrefWindow = app.window(title='Mseq Preferences')
                okPrefButton = mseqPrefWindow.child_window(title="&OK", class_name="Button")
                okPrefButton.click_input()
                if log_file:
                    enhanced_logging("Clicked OK on Preferences", log_file)
            else:
                if log_file:
                    enhanced_logging("Preferences window never appeared", log_file)
                return False
        except Exception as pref_error:
            if log_file:
                enhanced_logging(f"Error with Preferences window: {str(pref_error)}", log_file)
            return False
        
        # Wait for and handle Copy Files dialog
        if log_file:
            enhanced_logging("Waiting for Copy Files dialog", log_file)
            
        try:
            if timings.wait_until(timeout=10, retry_interval=0.5, func=lambda: IsCopyFilesOpen(app), value=True):
                copySeqFilesWindow = app.window(title='Copy sequence files', class_name='#32770')
                
                # Try to select files using various methods
                files_selected = False
                
                # Method 1: Original ShellView approach
                try:
                    if log_file:
                        enhanced_logging("Trying ShellView method for file selection", log_file)
                    shellViewFiles = copySeqFilesWindow.child_window(title="ShellView", class_name="SHELLDLL_DefView")
                    listView = shellViewFiles.child_window(class_name="DirectUIHWND")
                    listView.click_input()
                    TimeSleep(0.3)
                    send_keys('^a')
                    files_selected = True
                    if log_file:
                        enhanced_logging("ShellView file selection successful", log_file)
                except Exception as shell_error:
                    if log_file:
                        enhanced_logging(f"ShellView error: {str(shell_error)}", log_file)
                    
                    # Method 2: Try using SysListView32
                    try:
                        if log_file:
                            enhanced_logging("Trying SysListView method", log_file)
                        list_view = copySeqFilesWindow.child_window(class_name="SysListView32")
                        list_view.click_input()
                        TimeSleep(0.3)
                        send_keys('^a')
                        files_selected = True
                        if log_file:
                            enhanced_logging("SysListView file selection successful", log_file)
                    except Exception as list_error:
                        if log_file:
                            enhanced_logging(f"SysListView error: {str(list_error)}", log_file)
                        
                        # Method 3: Just focus window and send Ctrl+A
                        try:
                            if log_file:
                                enhanced_logging("Trying window focus method", log_file)
                            copySeqFilesWindow.set_focus()
                            TimeSleep(0.5)
                            send_keys('^a')
                            files_selected = True
                            if log_file:
                                enhanced_logging("Window focus file selection successful", log_file)
                        except Exception as focus_error:
                            if log_file:
                                enhanced_logging(f"Window focus error: {str(focus_error)}", log_file)
                
                if not files_selected:
                    if log_file:
                        enhanced_logging("All file selection methods failed", log_file)
                    return False
                
                # Click Open button
                try:
                    if log_file:
                        enhanced_logging("Clicking Open button", log_file)
                    copySeqOpenButton = copySeqFilesWindow.child_window(title="&Open", class_name="Button")
                    copySeqOpenButton.click_input()
                except Exception as open_error:
                    if log_file:
                        enhanced_logging(f"Error clicking Open: {str(open_error)}", log_file)
                    # Try alternative Open buttons
                    open_clicked = False
                    for title in ["&Open", "Open"]:
                        try:
                            open_btn = copySeqFilesWindow.child_window(title=title, class_name="Button")
                            open_btn.click_input()
                            open_clicked = True
                            if log_file:
                                enhanced_logging(f"Clicked {title} button", log_file)
                            break
                        except:
                            pass
                            
                    if not open_clicked:
                        if log_file:
                            enhanced_logging("Failed to click Open button", log_file)
                        return False
            else:
                if log_file:
                    enhanced_logging("Copy Files dialog never appeared", log_file)
                return False
        except Exception as copy_error:
            if log_file:
                enhanced_logging(f"Copy Files dialog error: {str(copy_error)}", log_file)
            return False
        
        # Handle remaining dialogs
        
        # Step 1: Handle File Error dialog
        if log_file:
            enhanced_logging("Waiting for File Error dialog", log_file)
            
        try:
            if timings.wait_until(timeout=15, retry_interval=0.5, func=lambda: IsErrorWindowOpen(app), value=True):
                fileErrorWindow = app.window(title='File error')
                fileErrorOkButton = fileErrorWindow.child_window(class_name="Button")
                fileErrorOkButton.click_input()
                if log_file:
                    enhanced_logging("Clicked OK on File Error dialog", log_file)
            else:
                if log_file:
                    enhanced_logging("File Error dialog did not appear (this may be normal)", log_file)
        except Exception as error_dialog_error:
            if log_file:
                enhanced_logging(f"Error handling File Error dialog: {str(error_dialog_error)}", log_file)
        
        # Step 2: Handle Call Bases dialog
        if log_file:
            enhanced_logging("Waiting for Call Bases dialog", log_file)
            
        try:
            if timings.wait_until(timeout=15, retry_interval=0.5, func=lambda: IsCallBasesOpen(app), value=True):
                callBasesWindow = app.window(title='Call bases?')
                callBasesYesButton = callBasesWindow.child_window(title="&Yes", class_name="Button")
                callBasesYesButton.click_input()
                if log_file:
                    enhanced_logging("Clicked Yes on Call Bases dialog", log_file)
            else:
                if log_file:
                    enhanced_logging("Call Bases dialog never appeared", log_file)
                return False
        except Exception as call_bases_error:
            if log_file:
                enhanced_logging(f"Error with Call Bases dialog: {str(call_bases_error)}", log_file)
            return False
        
        # Step 3: Wait for processing to complete
        if log_file:
            enhanced_logging("Waiting for processing to complete", log_file)
            
        try:
            if timings.wait_until(timeout=60, retry_interval=1.0, func=lambda: FivetxtORlowQualityWindow(app, orderpath), value=True):
                if log_file:
                    enhanced_logging("Processing completed", log_file)
            else:
                if log_file:
                    enhanced_logging("Processing timed out - checking for success anyway", log_file)
                # Double-check if we have 5txt files despite timeout
                if CheckFor5txt(orderpath):
                    if log_file:
                        enhanced_logging("Found 5txt files despite timeout - continuing", log_file)
                else:
                    if log_file:
                        enhanced_logging("No 5txt files found - processing failed", log_file)
                    return False
        except Exception as processing_error:
            if log_file:
                enhanced_logging(f"Error waiting for processing: {str(processing_error)}", log_file)
            return False
        
        # Step 4: Handle Low Quality Files dialog if it appears
        if app.window(title="Low quality files skipped").exists():
            if log_file:
                enhanced_logging("Handling Low Quality Files dialog", log_file)
            lowQualityWindow = app.window(title='Low quality files skipped')
            lowQualityOkButton = lowQualityWindow.child_window(class_name="Button")
            lowQualityOkButton.click_input()
        
        # Step 5: Close Read Info window if it appears
        if log_file:
            enhanced_logging("Checking for Read Info window", log_file)
            
        TimeSleep(1.0)  # Give time for Read Info to appear
        if app.window(title_re='Read information for*').exists():
            if log_file:
                enhanced_logging("Closing Read Info window", log_file)
            readWindow = app.window(title_re='Read information for*')
            readWindow.close()
        
        if log_file:
            enhanced_logging(f"Successfully completed mSeq for {orderpath}", log_file)
        return True
        
    except Exception as e:
        if log_file:
            enhanced_logging(f"Critical error in MseqOrder: {str(e)}", log_file)
            if isinstance(e, Exception):
                enhanced_logging(f"Error details: {traceback.format_exc()}", log_file)
        
        # Try to recover by pressing Escape
        try:
            for _ in range(3):
                send_keys('{ESC}')
                TimeSleep(0.5)
        except:
            pass
            
        return False

def GetMseq():
    """Function returns both the app object and its main window."""
    from pywinauto.findwindows import ElementNotFoundError, ElementAmbiguousError
    from pywinauto import Application, timings
    
    try:
        app = Application(backend='win32').connect(title_re='Mseq*', timeout=1)
    except (ElementNotFoundError, timings.TimeoutError):
        try:
            app = Application(backend='win32').connect(title_re='mSeq*', timeout=1)
        except (ElementNotFoundError, timings.TimeoutError):
            app = Application(backend='win32').start('cmd /c "cd /d C:\\DNA\\Mseq4\\bin && C:\\DNA\\Mseq4\\bin\\j.exe -jprofile mseq.ijl"', wait_for_idle=False).connect(title='mSeq', timeout=10)
        except ElementAmbiguousError:
            app = Application(backend='win32').connect(title_re='mSeq*', found_index=0, timeout=1)
            app2 = Application(backend='win32').connect(title_re='mSeq*', found_index=1, timeout=1)
            app2.kill()
    except ElementAmbiguousError:
        app = Application(backend='win32').connect(title_re='Mseq*', found_index=0, timeout=1)
        app2 = Application(backend='win32').connect(title_re='Mseq*', found_index=1, timeout=1)
        app2.kill()
    
    if app.window(title_re='mSeq*').exists()==False:
        mainWindow = app.window(title_re='Mseq*')
    else:
        mainWindow = app.window(title_re='mSeq*')

    return app, mainWindow

def IsBrowseDialogOpen(myapp):
    return myapp.window(title='Browse For Folder').exists()

def IsPreferencesOpen(myapp):
    return myapp.window(title='Mseq Preferences').exists()

def IsCopyFilesOpen(myapp):
    return myapp.window(title='Copy sequence files').exists()

def IsErrorWindowOpen(myapp):
    return myapp.window(title='File error').exists()

def IsCallBasesOpen(myapp):
    return myapp.window(title='Call bases?').exists()

def IsReadInfoOpen(myapp):
    return myapp.window(title_re='Read information for*').exists()

def FivetxtORlowQualityWindow(myApp, orderpath):
    if myApp.window(title="Low quality files skipped").exists():
        return True
    if CheckFor5txt(orderpath):
        return True
    return False

#############################################
# FOLDER NAVIGATION FUNCTIONS
#############################################

def NavigateToFolder_Win10(browsefolderswindow, path, log_file=None):
    """Original navigation function that works on Windows 10"""
    from time import sleep as TimeSleep

    if log_file:
        enhanced_logging(f"Using Windows 10 navigation for path: {path}", log_file)
    
    browsefolderswindow.set_focus()
    treeView = browsefolderswindow.child_window(class_name="SysTreeView32")
    desktopPath = f'\\Desktop\\{get_virtual_folder_name()}'
    
    try:
        item = treeView.get_item(desktopPath)
        data = path.split('\\')
        for folder in data:
            if 'P:' in folder:
                folder = sub('P:', r'ABISync (P:)', folder)
            if 'H:' in folder:
                folder = sub('H:', r'Tyler (\\\\w2k16\\users) (H:)', folder)
                
            folder_found = False
            for child in item.children():
                if child.text() == folder:
                    if log_file:
                        enhanced_logging(f"Found folder: {folder}", log_file)
                    browsefolderswindow.set_focus()
                    item = child
                    item.click_input()
                    folder_found = True
                    TimeSleep(0.2)  # Short delay after each click
                    break
                    
            if not folder_found and log_file:
                enhanced_logging(f"Could not find folder: {folder}", log_file)
                
        # Final folder should now be selected
        return True
        
    except Exception as e:
        if log_file:
            enhanced_logging(f"Windows 10 navigation error: {str(e)}", log_file)
        return False

def NavigateToFolder(browsefolderswindow, path, log_file=None):
    """Master function that selects the appropriate navigation method based on OS version"""
    windows_version = get_windows_version()
    
    if log_file:
        enhanced_logging(f"Detected OS: {windows_version}", log_file)
        enhanced_logging(f"Attempting to navigate to: {path}", log_file)
    
    if windows_version == "Windows 11":
        return NavigateToFolder_Win11(browsefolderswindow, path, log_file)
    else:
        return NavigateToFolder_Win10(browsefolderswindow, path, log_file)
    
###Old version of NavigateToFolder_Win11 below
'''
def NavigateToFolder_Win11(browsefolderswindow, path, log_file=None):
    """Enhanced navigation function for Windows 11 using Tab navigation and targeted element finding"""
    from pywinauto.keyboard import send_keys
    from time import sleep as TimeSleep
    import os
    
    if log_file:
        enhanced_logging(f"Using Windows 11 navigation for path: {path}", log_file)
        enhanced_logging(f"Target folder name: {os.path.basename(path)}", log_file)
    
    # Step 1: Reset focus and navigate to the tree view
    try:
        browsefolderswindow.set_focus()
        TimeSleep(0.5)
        
        # Use Tab to move focus to the tree view (usually takes 2-3 tabs)
        if log_file:
            enhanced_logging("Resetting focus and navigating to tree view with Tab", log_file)
        
        # Start with Tab to ensure we're in the dialog controls
        send_keys('{TAB}')
        TimeSleep(0.3)
        send_keys('{TAB}')
        TimeSleep(0.3)
        
        
        if log_file:
            enhanced_logging("Navigated to tree view, looking for This PC", log_file)
        
        # Step 2: Now find and navigate to This PC first
        all_elements = browsefolderswindow.descendants()
        this_pc_found = False
        
        for element in all_elements:
            try:
                if element.is_visible() and "This PC" in element.window_text():
                    if log_file:
                        enhanced_logging(f"Found This PC element: {element.window_text()}", log_file)
                    element.click_input()
                    this_pc_found = True
                    TimeSleep(0.8)  # Wait for expansion
                    break
            except Exception as e:
                if log_file:
                    enhanced_logging(f"Error checking element: {str(e)}", log_file)
        
        if not this_pc_found:
            if log_file:
                enhanced_logging("Could not find This PC element, trying keyboard navigation", log_file)
            # Try keyboard navigation if clicking fails
            send_keys('{HOME}')  # Move to top of tree
            TimeSleep(0.3)
            
            # Search for This PC using arrow keys
            for _ in range(15):  # Try reasonable number of items
                # Check if current selection is This PC
                try:
                    focused_element = browsefolderswindow.get_focus()
                    if "This PC" in focused_element.window_text():
                        if log_file:
                            enhanced_logging("Found This PC with keyboard navigation", log_file)
                        this_pc_found = True
                        break
                except:
                    pass
                send_keys('{DOWN}')
                TimeSleep(0.2)
        
        if not this_pc_found:
            if log_file:
                enhanced_logging("Failed to find This PC node", log_file)
            return False
        
        # Expand This PC if needed
        send_keys('{RIGHT}')
        TimeSleep(0.5)
        
        # Step 3: Find and click on ABISync (P:)
        abisync_found = False
        all_elements = browsefolderswindow.descendants()  # Refresh elements
        
        for element in all_elements:
            try:
                element_text = element.window_text()
                if element.is_visible() and ("ABISync (P:)" in element_text or "ABISync" in element_text):
                    if log_file:
                        enhanced_logging(f"Found P: drive element: {element_text}", log_file)
                    element.click_input()
                    abisync_found = True
                    TimeSleep(0.8)
                    break
            except Exception as e:
                if log_file:
                    enhanced_logging(f"Error checking P: drive element: {str(e)}", log_file)
        
        if not abisync_found:
            if log_file:
                enhanced_logging("Could not find ABISync (P:) element, trying keyboard navigation", log_file)
            # Try keyboard navigation
            for _ in range(20):  # Try reasonable number of items
                try:
                    focused_element = browsefolderswindow.get_focus()
                    text = focused_element.window_text()
                    if "ABISync" in text or "(P:)" in text:
                        if log_file:
                            enhanced_logging(f"Found P: drive with keyboard navigation: {text}", log_file)
                        abisync_found = True
                        break
                except:
                    pass
                send_keys('{DOWN}')
                TimeSleep(0.2)
        
        if not abisync_found:
            if log_file:
                enhanced_logging("Failed to find ABISync (P:) drive", log_file)
            return False
        
        # Expand P: drive
        send_keys('{RIGHT}')
        TimeSleep(0.5)
        
        # Step 4: Navigate the rest of the path using the text elements
        path_parts = path.split('\\')
        # Skip the drive part
        path_parts = [part for part in path_parts if "P:" not in part and part]
        
        if log_file:
            enhanced_logging(f"Navigating remaining path parts: {path_parts}", log_file)
        
        for part in path_parts:
            part_found = False
            all_elements = browsefolderswindow.descendants()  # Refresh elements
            
            # Try direct click first
            for element in all_elements:
                try:
                    if element.is_visible() and part in element.window_text():
                        if log_file:
                            enhanced_logging(f"Found path component: {part}", log_file)
                        element.click_input()
                        part_found = True
                        TimeSleep(0.5)
                        
                        # If this isn't the last part, expand it
                        if part != path_parts[-1]:
                            send_keys('{RIGHT}')
                            TimeSleep(0.5)
                        break
                except Exception as e:
                    if log_file:
                        enhanced_logging(f"Error checking path component: {str(e)}", log_file)
            
            # If direct click failed, try keyboard navigation
            if not part_found:
                if log_file:
                    enhanced_logging(f"Could not directly click {part}, trying keyboard navigation", log_file)
                
                # Start scanning from current position
                for _ in range(30):  # Try reasonable number of items
                    try:
                        focused_element = browsefolderswindow.get_focus()
                        if part in focused_element.window_text():
                            if log_file:
                                enhanced_logging(f"Found {part} with keyboard navigation", log_file)
                            part_found = True
                            
                            # If this isn't the last part, expand it
                            if part != path_parts[-1]:
                                send_keys('{RIGHT}')
                                TimeSleep(0.5)
                            break
                    except:
                        pass
                    send_keys('{DOWN}')
                    TimeSleep(0.2)
            
            if not part_found:
                if log_file:
                    enhanced_logging(f"Failed to find path component: {part}", log_file)
                return False
        
        # Final check - verify we've reached the target folder
        try:
            ok_button = browsefolderswindow.child_window(title="OK", class_name="Button")
            if ok_button.exists() and ok_button.is_enabled():
                if log_file:
                    enhanced_logging("Successfully navigated to target folder", log_file)
                return True
        except:
            pass
        
        return True
        
    except Exception as e:
        if log_file:
            enhanced_logging(f"Critical error in Win11 navigation: {str(e)}", log_file)
        return False
'''
def NavigateToFolder_Win11(browsefolderswindow, path, log_file=None):
    """Enhanced tree-focused navigation function for Windows 11 that also works on Windows 10"""
    from time import sleep as TimeSleep
    import os

    if log_file:
        enhanced_logging(f"Using Windows 11 navigation for path: {path}", log_file)
        enhanced_logging(f"Target folder name: {os.path.basename(path)}", log_file)

    try:
        # Establish focus on the dialog
        browsefolderswindow.set_focus()
        TimeSleep(0.5)
        
        # Try to find the tree view control
        try:
            treeView = browsefolderswindow.child_window(class_name="SysTreeView32")
            if log_file:
                enhanced_logging("Found SysTreeView32 control", log_file)
        except Exception as e:
            if log_file:
                enhanced_logging(f"Could not find tree view: {str(e)}", log_file)
            return False
        
        # Try different starting points in the tree
        desktop_paths = [
            f'\\Desktop\\{get_virtual_folder_name()}',
            '\\Desktop\\This PC',
            '\\This PC',
            '\\Computer'
        ]
        
        item = None
        for desktop_path in desktop_paths:
            try:
                if log_file:
                    enhanced_logging(f"Trying to find tree node: {desktop_path}", log_file)
                item = treeView.get_item(desktop_path)
                if log_file:
                    enhanced_logging(f"Successfully found tree node: {desktop_path}", log_file)
                break
            except Exception as e:
                if log_file:
                    enhanced_logging(f"Could not find tree node {desktop_path}: {str(e)}", log_file)
        
        # If we found a starting point, navigate through the path
        if item:
            data = path.split('\\')
            for folder in data:
                if 'P:' in folder:
                    folder = sub('P:', r'ABISync (P:)', folder)
                if 'H:' in folder:
                    folder = sub('H:', r'Tyler (\\\\w2k16\\users) (H:)', folder)
                
                folder_found = False
                children = item.children()
                if log_file:
                    child_texts = [child.text() for child in children]
                    enhanced_logging(f"Available children: {child_texts}", log_file)
                
                for child in children:
                    if child.text() == folder:
                        if log_file:
                            enhanced_logging(f"Found folder: {folder}", log_file)
                        browsefolderswindow.set_focus()
                        item = child
                        item.click_input()
                        folder_found = True
                        TimeSleep(0.5)  # Increased delay for reliable tree navigation
                        break
                
                if not folder_found:
                    if log_file:
                        enhanced_logging(f"Could not find folder: {folder}", log_file)
                    
                    # Try case-insensitive partial match as fallback
                    for child in children:
                        if folder.lower() in child.text().lower():
                            if log_file:
                                enhanced_logging(f"Found partial match: {child.text()} for {folder}", log_file)
                            browsefolderswindow.set_focus()
                            item = child
                            item.click_input()
                            folder_found = True
                            TimeSleep(0.5)
                            break
                
                if not folder_found:
                    if log_file:
                        enhanced_logging(f"Failed to find folder {folder} - stopping navigation", log_file)
                    return False
            
            # Navigation completed successfully
            if log_file:
                enhanced_logging("Tree navigation successful", log_file)
            return True
        
        if log_file:
            enhanced_logging("Could not find starting point in tree view", log_file)
        return False
        
    except Exception as e:
        if log_file:
            enhanced_logging(f"Critical error in navigation: {str(e)}", log_file)
            if isinstance(e, Exception):
                import traceback
                enhanced_logging(f"Error details: {traceback.format_exc()}", log_file)
        return False
#############################################
# MAIN CODE
#############################################

if __name__ == "__main__":
    # Initialize log file at program start
    log_path, log_file = init_log_file()
    
    try:
        log_message(log_file, "Starting program")
        windows_version = get_windows_version()
        log_message(log_file, f"Running on {windows_version}")
        
        # Run batch file to generate key file
        batPath = "P:\\generate-data-sorting-key-file.bat"
        try:
            SubprocessRun(batPath, shell=True, check=True)
            log_message(log_file, "Successfully generated data sorting key file")
        except CalledProcessError:
            error = f"Batch file {batPath} failed to execute"
            log_message(log_file, error)
            print(error)
            
        keyFilePathName = 'P:\\order_key.txt'
        reinjectListFileName = f'Reinject List_{GetFormattedDate()}.xlsx'
        reinjectListFilePath = f'P:\\Data\\Reinjects\\{reinjectListFileName}'
        completeReinjectList = []
        reinjectPrepList = []
        completeRawReinjectList = []
        
        try:
            key = loadtxt(keyFilePathName, dtype=str, delimiter='\t')
            AdjustFullKeyToABIChars
            log_message(log_file, f"Successfully loaded key file with {len(key)} entries")
        except Exception as e:
            error = f"Error loading key file: {str(e)}"
            log_message(log_file, error)
            print(error)
            raise

        # Get folder selection from user
        dataFolderPath = filedialog.askdirectory(title="Select today's data folder. To mseq orders.")
        if not dataFolderPath:
            log_message(log_file, "User canceled folder selection")
            if log_file and not log_file.closed:
                log_file.close()
            sys.exit(0)
            
        dataFolderPath = sub(r'/','\\\\', dataFolderPath)
        log_message(log_file, f"Selected folder: {dataFolderPath}")
        
        # Get folders to process
        bioIFolders = GetInumberFolders(dataFolderPath)
        immediateOrderFolders = GetImmediateOrders(dataFolderPath)
        pcrFolders = GetPCRFolders(dataFolderPath)
        log_message(log_file, f"Found {len(bioIFolders)} BioI folders, {len(immediateOrderFolders)} immediate order folders, {len(pcrFolders)} PCR folders")
        
        # Import required modules
        from pywinauto import Application
        from pywinauto.findwindows import ElementNotFoundError, ElementAmbiguousError
        from pywinauto.keyboard import send_keys
        from pywinauto import timings
        
        # Process BioI folders
        log_message(log_file, "=== Processing BioI Folders ===")
        for bioIFolder in bioIFolders:
            log_message(log_file, f"Checking BioI folder: {OsPath.basename(bioIFolder)}")
            orderFolders = GetOrderFolders(bioIFolder)
            
            for orderFolder in orderFolders:
                try:
                    orderName = OsPath.basename(orderFolder)
                    log_message(log_file, f"Processing order: {orderName}")
                    
                    if 'andreev' in orderFolder.lower():
                        log_message(log_file, f"Skipping mSeq for Andreev order: {orderName}")
                        # For Dmitri orders, just check for completion
                        orderNumber = GetOrderNumberFromFolderName(orderFolder)
                        orderKey = GetOrderList(orderNumber)
                        allab1Files = GetAllab1Files(orderFolder)
                        
                        if len(orderKey) != len(allab1Files):
                            if not OsPath.exists(f'{dataFolderPath}\\IND Not Ready'):
                                mkdir(f'{dataFolderPath}\\IND Not Ready')
                            destination = f'{dataFolderPath}\\IND Not Ready\\'
                            ShutilMove(orderFolder, destination)
                            log_message(log_file, f"Moved incomplete order to IND Not Ready: {orderName}")
                        continue
                        
                    # Process normal orders
                    orderNumber = GetOrderNumberFromFolderName(orderFolder)
                    orderKey = GetOrderList(orderNumber)
                    allab1Files = GetAllab1Files(orderFolder)
                    mseqed, braces, yesab1 = CheckOrder(orderFolder)
                    
                    log_message(log_file, f"Order status - mSeqed: {mseqed}, has braces: {braces}, has ab1 files: {yesab1}")
                    log_message(log_file, f"File counts - Expected: {len(orderKey)}, Found: {len(allab1Files)}")
                    
                    if not mseqed and not braces:
                        if len(orderKey) == len(allab1Files) and yesab1:
                            result = MseqOrder(orderFolder, log_file)
                            if result:
                                log_message(log_file, f"mSeq completed successfully: {orderName}")
                            else:
                                log_message(log_file, f"mSeq failed: {orderName}")
                        else:
                            if not OsPath.exists(f'{dataFolderPath}\\IND Not Ready'):
                                mkdir(f'{dataFolderPath}\\IND Not Ready')
                            destination = f'{dataFolderPath}\\IND Not Ready\\'
                            ShutilMove(orderFolder, destination)
                            log_message(log_file, f"Moved incomplete order to IND Not Ready: {orderName}")
                            
                except Exception as e:
                    error_details = traceback.format_exc()
                    log_message(log_file, f"ERROR processing {OsPath.basename(orderFolder)}: {str(e)}")
                    log_message(log_file, error_details)
                    print(f"Error processing {OsPath.basename(orderFolder)}: {str(e)}")
        
        # Process immediate order folders
        log_message(log_file, "=== Processing Immediate Order Folders ===")
        INDNotReadyIsTheSelectedFolder = (OsPath.basename(dataFolderPath) == 'IND Not Ready')
        log_message(log_file, f"Selected folder is 'IND Not Ready': {INDNotReadyIsTheSelectedFolder}")
        
        for orderFolder in immediateOrderFolders:
            try:
                orderName = OsPath.basename(orderFolder)
                log_message(log_file, f"Processing immediate order: {orderName}")
                
                if 'andreev' in orderFolder.lower():
                    # Handle Andreev orders
                    if INDNotReadyIsTheSelectedFolder:
                        orderNumber = GetOrderNumberFromFolderName(orderFolder)
                        orderKey = GetOrderList(orderNumber)
                        allab1Files = GetAllab1Files(orderFolder)
                        
                        if len(orderKey) == len(allab1Files):
                            destination = GetDestination(orderName, dataFolderPath)
                            ShutilMove(orderFolder, destination)
                            log_message(log_file, f"Moved complete Andreev order to destination: {orderName}")
                    continue
                
                # Process normal orders
                orderNumber = GetOrderNumberFromFolderName(orderFolder)
                orderKey = GetOrderList(orderNumber)
                allab1Files = GetAllab1Files(orderFolder)
                mseqed, braces, yesab1 = CheckOrder(orderFolder)
                
                log_message(log_file, f"Order status - mSeqed: {mseqed}, has braces: {braces}, has ab1 files: {yesab1}")
                log_message(log_file, f"File counts - Expected: {len(orderKey)}, Found: {len(allab1Files)}")
                
                if not mseqed and not braces:
                    if len(orderKey) == len(allab1Files) and yesab1:
                        result = MseqOrder(orderFolder, log_file)
                        if result:
                            log_message(log_file, f"mSeq completed successfully: {orderName}")
                        else:
                            log_message(log_file, f"mSeq failed: {orderName}")
                            
                        if INDNotReadyIsTheSelectedFolder:
                            destination = GetDestination(orderName, dataFolderPath)
                            app, mseqMainWindow = GetMseq()
                            mseqMainWindow = None
                            app.kill()
                            TimeSleep(.1)
                            ShutilMove(orderFolder, destination)
                            log_message(log_file, f"Moved completed order to destination: {orderName}")
                    else:
                        if not INDNotReadyIsTheSelectedFolder:
                            if not OsPath.exists(f'{dataFolderPath}\\IND Not Ready'):
                                mkdir(f'{dataFolderPath}\\IND Not Ready')
                            destination = f'{dataFolderPath}\\IND Not Ready\\'
                            ShutilMove(orderFolder, destination)
                            log_message(log_file, f"Moved incomplete order to IND Not Ready: {orderName}")
                
                elif mseqed and not braces and INDNotReadyIsTheSelectedFolder:
                    destination = GetDestination(orderName, dataFolderPath)
                    app, mseqMainWindow = GetMseq()
                    mseqMainWindow = None
                    app.kill()
                    TimeSleep(.1)
                    ShutilMove(orderFolder, destination)
                    log_message(log_file, f"Moved previously mSeqed order to destination: {orderName}")
                    
            except Exception as e:
                error_details = traceback.format_exc()
                log_message(log_file, f"ERROR processing immediate order {OsPath.basename(orderFolder)}: {str(e)}")
                log_message(log_file, error_details)
                print(f"Error processing immediate order {OsPath.basename(orderFolder)}: {str(e)}")
        
        # Process PCR folders
        log_message(log_file, "=== Processing PCR Folders ===")
        for pcrFolder in pcrFolders:
            try:
                folder_name = OsPath.basename(pcrFolder)
                log_message(log_file, f"Processing PCR folder: {folder_name}")
                
                mseqed, braces, yesab1 = CheckOrder(pcrFolder)
                log_message(log_file, f"PCR folder status - mSeqed: {mseqed}, has braces: {braces}, has ab1 files: {yesab1}")
                
                if not mseqed and not braces and yesab1:
                    result = MseqOrder(pcrFolder, log_file)
                    if result:
                        log_message(log_file, f"mSeq completed successfully: {folder_name}")
                    else:
                        log_message(log_file, f"mSeq failed: {folder_name}")
                else:
                    log_message(log_file, f"mSeq not performed: {folder_name}")
                    
            except Exception as e:
                error_details = traceback.format_exc()
                log_message(log_file, f"ERROR processing PCR folder {OsPath.basename(pcrFolder)}: {str(e)}")
                log_message(log_file, error_details)
                print(f"Error processing PCR folder {OsPath.basename(pcrFolder)}: {str(e)}")
        
        # Close mSeq application
        try:
            app, mseqMainWindow = GetMseq()
            mseqMainWindow = None
            app.kill()
            log_message(log_file, "Successfully closed mSeq application")
        except Exception as e:
            log_message(log_file, f"Error closing mSeq: {str(e)}")
        
        log_message(log_file, "Processing completed successfully")
        print("\nALL DONE")
        
    except Exception as e:
        error_details = traceback.format_exc()
        log_message(log_file, f"CRITICAL ERROR: {str(e)}")
        log_message(log_file, error_details)
        print(f"Critical error: {str(e)}")
        
    finally:
        # Close the log file
        if log_file and not log_file.closed:
            log_message(log_file, "Program execution ended")
            log_file.close()