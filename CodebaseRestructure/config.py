# config.py
class MseqConfig:
    # Paths
    PYTHON32_PATH = r"C:\Python312-32\python.exe"
    MSEQ_PATH = r"C:\DNA\Mseq4\bin"
    MSEQ_EXECUTABLE = r"j.exe -jprofile mseq.ijl"
    
    # Network drives
    NETWORK_DRIVES = {
        "P:": r"ABISync (P:)",
        "H:": r"Tyler (\\w2k16\users) (H:)"
    }
    
    # Timeouts for UI operations
    TIMEOUTS = {
        "browse_dialog": 5,
        "preferences": 5, 
        "copy_files": 5,
        "error_window": 20,
        "call_bases": 10,
        "process_completion": 45,
        "read_info": 5
    }
    
    # Special folders
    IND_NOT_READY_FOLDER = "IND Not Ready"
    
    # Mseq artifacts
    MSEQ_ARTIFACTS = {'chromat_dir', 'edit_dir', 'phd_dir', 'mseq4.ini'}
    
    # File types
    TEXT_FILES = [
        '.raw.qual.txt',
        '.raw.seq.txt',
        '.seq.info.txt',
        '.seq.qual.txt',
        '.seq.txt'
    ]