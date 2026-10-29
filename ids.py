from pywinauto.application import Application
from pywinauto import Desktop
import time

connect_launcher = r"C:\Users\RPA_DEVWNS02\RDP6\ConnectLauncher.exe"
connect_file = r"C:\Users\RPA_DEVWNS02\Desktop\SI-Reports.connect"

username = "RPA_DEVWNS02"
password = "Password2026"

def handle_security_prompt():
    try:
        # 1. Connect to the global desktop using UIA backend
        # This is necessary because Windows Security is a system-level process
        print("Searching for Windows Security dialog...")
        
        # We search for the window specifically titled "Windows Security"
        sec_dlg = Desktop(backend="uia").window(title="Windows Security", visible_only=True)
        
        if sec_dlg.exists(timeout=10):
            print("Found security prompt. Bringing to focus.")
            sec_dlg.set_focus()
            
            # 2. Identify the Password field. 
            # In these dialogs, the password box is usually the first 'Edit' control
            # You can also use auto_id="PasswordField" or similar if found in dump_tree
            pw_field = sec_dlg.child_window(control_type="Edit", found_index=0)
            
            # 3. Enter the password
            # Using {ENTER} at the end to submit the form immediately
            print("Typing password...")
            pw_field.type_keys(password + "{ENTER}", protect_first=True)
            
            print("Login submitted.")
        else:
            print("Security dialog did not appear.")

    except Exception as e:
        print(f"Failed to interact with security dialog: {e}")


def launch_and_login():
    # 1. Start the launcher
    Application().start(f'"{connect_launcher}" "{connect_file}"')
    print("Waiting for Remote Access window...")
    time.sleep(6) 
    
    try:
        # 2. Use Desktop with UIA backend (matches your log output)
        # This is more robust for RDP windows that jump processes
        dlg = Desktop(backend="uia").window(title="Remote Access", visible_only=True)
        dlg.set_focus()

        # 3. Use the exact auto_id and control_type from your log
        # Username field: auto_id="5"
        # Password field: auto_id="6"
        # Connect button: auto_id="8"
        
        print("Entering credentials...")
        
        # Using type_keys is safer for RDP fields than set_edit_text
        dlg.child_window(auto_id="5", control_type="Edit").type_keys(username, with_spaces=True)
        dlg.child_window(auto_id="6", control_type="Edit").type_keys("^a{BACKSPACE}" + password, with_spaces=True)
        # 4. Click the Connect button
        dlg.child_window(auto_id="8", control_type="Button").click()

        print("Login automation complete.")
        
    except Exception as e:
        print(f"Error details: {e}")

if __name__ == "__main__":
    launch_and_login()
    handle_security_prompt()