import pyautogui
import time

# Give yourself a moment to switch to the window
time.sleep(2)

try:
    # Locate the center of the password box on screen
    location = pyautogui.locateOnScreen('password_field.png', confidence=0.8)
    
    if location:
        # Click the box to focus
        pyautogui.click(location)
        pyautogui.write('Password2026', interval=0.1)
        # Type the password and press enter
        pyautogui.press('enter')
    else:
        print("Could not find the password field on screen.")
except Exception as e:
    print(f"An error occurred: {e}")