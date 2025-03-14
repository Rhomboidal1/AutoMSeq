import os
import time
from pywinauto import Desktop
from pywinauto.application import Application

def dialog_explorer():
    """Wait for new dialog windows to appear and capture their details"""
    print("Dialog Explorer started. Press Ctrl+C to stop.")
    print("Open dialogs in mSeq to capture their details...")
    
    # Create logs directory
    os.makedirs("dialog_logs", exist_ok=True)
    
    # Get initial set of windows
    initial_windows = set(window.window_text() for window in Desktop(backend="win32").windows())
    
    try:
        while True:
            # Get current windows
            current_windows = set(window.window_text() for window in Desktop(backend="win32").windows())
            
            # Find new windows
            new_windows = current_windows - initial_windows
            if new_windows:
                print(f"New windows detected: {new_windows}")
                
                # Capture details of new windows
                for window_text in new_windows:
                    if window_text:  # Skip empty titles
                        try:
                            # Find the window by its text
                            window = Desktop(backend="win32").window(title=window_text)
                            
                            # Create a safe filename
                            safe_title = "".join(c for c in window_text if c.isalnum() or c in " _-").strip()
                            if not safe_title:
                                safe_title = "Untitled_Dialog"
                            
                            filename = f"{safe_title}_{time.strftime('%H%M%S')}.txt"
                            filepath = os.path.join("dialog_logs", filename)
                            
                            # Save window details
                            with open(filepath, 'w') as f:
                                f.write(f"Window Title: {window.window_text()}\n")
                                f.write(f"Window Class: {window.class_name()}\n")
                                f.write(f"Window Rectangle: {window.rectangle()}\n\n")
                                f.write("Controls:\n")
                                
                                # Get all controls
                                for control in window.descendants():
                                    try:
                                        f.write(f"- {control.window_text()} ({control.class_name()})\n")
                                    except:
                                        f.write(f"- [Error getting control details]\n")
                            
                            print(f"Saved details to {filepath}")
                        except Exception as e:
                            print(f"Error capturing window '{window_text}': {e}")
                
                # Update initial windows
                initial_windows = current_windows
            
            # Wait before checking again
            time.sleep(0.5)
    
    except KeyboardInterrupt:
        print("\nDialog Explorer stopped.")

if __name__ == "__main__":
    dialog_explorer()