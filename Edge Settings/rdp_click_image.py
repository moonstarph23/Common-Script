import sys
import time
from typing import Optional
import numpy as np
import win32gui
import win32con

try:
    import pyautogui
    import cv2
    from PIL import Image
    import mss
except ImportError as e:
    print(f"Missing required package: {e}")
    pyautogui = None
    cv2 = None
    mss = None


def minimize_cmd_windows():
    """Find and minimize all Command Prompt (cmd) and OpenConsole windows"""
    try:
        windows_found = 0
        errors = []

        def window_enum_callback(hwnd, wildcard):
            nonlocal windows_found
            try:
                if win32gui.IsWindowVisible(hwnd):
                    window_text = win32gui.GetWindowText(hwnd)
                    class_name = win32gui.GetClassName(hwnd)
                    # Get process name to check for openconsole.exe
                    try:
                        import psutil
                        _, pid = win32gui.GetWindowThreadProcessId(hwnd)
                        process = psutil.Process(pid)
                        process_name = process.name().lower()
                        process_exe = process.exe().lower() if hasattr(process, 'exe') else ""
                    except:
                        process_name = ""
                        process_exe = ""
                    
                    # Check for various console window types
                    is_console = (
                        class_name == "ConsoleWindowClass" or 
                        class_name == "CASCADIA_HOSTING_WINDOW_CLASS" or  # Windows Terminal
                        "Command Prompt" in window_text or 
                        "cmd" in window_text.lower() or
                        "openconsole.exe" in process_name or
                        "openconsole" in process_exe or
                        "conhost.exe" in process_name or
                        "Windows Terminal" in window_text or
                        "PowerShell" in window_text or
                        # Check if window title contains common terminal indicators
                        ("python" in window_text.lower() and ("r:\\" in window_text.lower() or "c:\\" in window_text.lower()))
                    )
                    
                    if is_console:
                        print(f"Minimizing console window: '{window_text}' (Class: {class_name}, Process: {process_name})")
                        win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                        windows_found += 1
            except Exception as e:
                errors.append(f"Error minimizing window handle {hwnd}: {e}")

        try:
            win32gui.EnumWindows(lambda hwnd, _: window_enum_callback(hwnd, ""), None)
        except Exception as e:
            errors.append(f"Error during EnumWindows: {e}")

        if windows_found > 0:
            print(f"Successfully minimized {windows_found} console window(s)")
        else:
            print("No console windows found to minimize")

        if errors:
            print("Errors during minimization:")
            for err in errors:
                print(err)
            return False, errors
        return True, None
        if errors:
            print("Errors during minimization:")
            for err in errors:
                print(err)
            return False, errors
        return True, None
        if errors:
            print("Errors during minimization:")
            for err in errors:
                print(err)
            return False, errors
        return True, None
    except Exception as e:
        print(f"Error during CMD window minimization attempt: {e}")
        import traceback
        print(traceback.format_exc())
        return False, [str(e)]


def minimize_python_host_window():
    """Find and minimize UiPath Python Host window"""
    try:
        windows_found = 0
        errors = []

        def window_enum_callback(hwnd, wildcard):
            nonlocal windows_found
            try:
                if win32gui.IsWindowVisible(hwnd):
                    window_text = win32gui.GetWindowText(hwnd)
                    if wildcard in window_text:
                        print(f"Minimizing window: {window_text}")
                        win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                        windows_found += 1
            except Exception as e:
                errors.append(f"Error minimizing window handle {hwnd}: {e}")

        try:
            win32gui.EnumWindows(lambda hwnd, _: window_enum_callback(hwnd, "UiPath.Python.Host"), None)
        except Exception as e:
            errors.append(f"Error during EnumWindows: {e}")

        if windows_found > 0:
            print(f"Successfully minimized {windows_found} Python Host window(s)")
        else:
            print("No UiPath.Python.Host windows found to minimize")

        if errors:
            print("Errors during minimization:")
            for err in errors:
                print(err)
            return False, errors
        return True, None
    except Exception as e:
        print(f"Error during window minimization attempt: {e}")
        import traceback
        print(traceback.format_exc())
        return False, [str(e)]


def click_images(imagePath1: str, imagePath2: str, timeout: int) -> None:
    """
    Try to click on both images specified by imagePath1 and imagePath2 until timeout is reached.
    Will retry every 2 seconds, alternating between the two images. Uses multiple detection methods.
    Searches on both screens. Does not throw any error if images are not found.
    """
    imagePath3 = r"R:\Finance\Revenue Audit\HOTEL\City Ledger Report\Processes\laterButton.PNG"
    imagePath4 = r"R:\Finance\Revenue Audit\HOTEL\City Ledger Report\Processes\laterButton.PNG"
    # Wait 2 minutes before minimizing the Python host window
    print("Waiting 5 seconds before minimizing windows...")
    time.sleep(5)
    
    # Minimize Command Prompt windows
    print("Attempting to minimize Command Prompt windows...")
    minimize_cmd_windows()
    
    # Minimize Python host window
    print("Attempting to minimize Python host window...")
    minimize_python_host_window()

    if pyautogui is None:
        print("PyAutoGUI not available")
        return
    
    print(f"Starting to search for images: {imagePath1}, {imagePath2}, {imagePath3}, and {imagePath4}")
    print(f"Timeout: {timeout} seconds")
    print("Using multiple detection methods...")
    
    end_time = time.time() + timeout
    attempt = 1
    attemptCount = 5
    
    def try_pyautogui_method(image_path, image_name):
        """Try using PyAutoGUI's built-in method with exact color matching on all screens"""
        # Get all monitors
        import tkinter as tk
        root = tk.Tk()
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        root.destroy()
        
        try:
            # Method 1: Standard PyAutoGUI (should work on all screens)
            location = pyautogui.locateCenterOnScreen(image_path, confidence=0.9)
            if location:
                print(f"{image_name} found with PyAutoGUI (high confidence) at {location}! Clicking...")
                pyautogui.click(location)
                return True
        except:
            pass
        
        try:
            # Method 2: Medium confidence
            location = pyautogui.locateCenterOnScreen(image_path, confidence=0.8)
            if location:
                print(f"{image_name} found with PyAutoGUI (medium confidence) at {location}! Clicking...")
                pyautogui.click(location)
                return True
        except:
            pass
        
        return False
    
    def try_opencv_method(image_path, image_name):
        """Try using OpenCV with exact color matching on both screens"""
        if cv2 is None:
            return False
        
        try:
            # Method 1: Try with full desktop screenshot (all monitors)
            screenshot = pyautogui.screenshot()
            screenshot_cv = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            
            # Read template
            template = cv2.imread(image_path, cv2.IMREAD_COLOR)
            if template is None:
                return False
            
            # Use only TM_CCOEFF_NORMED with high threshold for exact matching
            result = cv2.matchTemplate(screenshot_cv, template, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
            
            # Only accept very high confidence matches (0.9+) for exact color matching
            if max_val > 0.9:
                # Calculate center
                h, w = template.shape[:2]
                center_x = max_loc[0] + w // 2
                center_y = max_loc[1] + h // 2
                
                print(f"{image_name} found with OpenCV (exact match) confidence {max_val:.3f} at ({center_x}, {center_y})! Clicking...")
                pyautogui.click(center_x, center_y)
                return True
            
            # Method 2: Try with MSS for individual monitor capture
            if mss is not None:
                with mss.mss() as sct:
                    monitors = sct.monitors[1:]  # Skip combined desktop
                    
                    for i, monitor in enumerate(monitors):
                        print(f"    Checking monitor {i+1} with OpenCV...")
                        
                        # Capture this specific monitor
                        screen_shot = sct.grab(monitor)
                        # Convert to PIL Image then to numpy array
                        screen_image = Image.frombytes("RGB", screen_shot.size, screen_shot.bgra, "raw", "BGRX")
                        screen_cv = cv2.cvtColor(np.array(screen_image), cv2.COLOR_RGB2BGR)
                        
                        # Template matching on this monitor
                        result = cv2.matchTemplate(screen_cv, template, cv2.TM_CCOEFF_NORMED)
                        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
                        
                        if max_val > 0.9:
                            # Calculate center with monitor offset
                            h, w = template.shape[:2]
                            center_x = max_loc[0] + w // 2 + monitor["left"]
                            center_y = max_loc[1] + h // 2 + monitor["top"]
                            
                            print(f"{image_name} found on monitor {i+1} with OpenCV confidence {max_val:.3f} at ({center_x}, {center_y})! Clicking...")
                            pyautogui.click(center_x, center_y)
                            return True
                
        except Exception as e:
            print(f"OpenCV error: {e}")
        
        return False
    
    def find_and_click_image(image_path, image_name):
        """Try multiple methods to find and click image"""
        print(f"  Trying PyAutoGUI methods for {image_name}...")
        if try_pyautogui_method(image_path, image_name):
            return True
        
        print(f"  Trying OpenCV methods for {image_name}...")
        if try_opencv_method(image_path, image_name):
            return True
        
        return False
    
    while time.time() < end_time and attempt <= attemptCount:
        clicked_any = False
        # Try first image
        print(f"Attempt {attempt}: Looking for first image...")
        if find_and_click_image(imagePath1, "first image"):
            clicked_any = True
        else:
            print(f"Attempt {attempt}: First image not found")
        
        # Try second image
        print(f"Attempt {attempt}: Looking for second image...")
        if find_and_click_image(imagePath2, "second image"):
            clicked_any = True
        else:
            print(f"Attempt {attempt}: Second image not found")
        
        # Try third image
        print(f"Attempt {attempt}: Looking for third image...")
        if find_and_click_image(imagePath3, "third image"):
            clicked_any = True
        else:
            print(f"Attempt {attempt}: Third image not found")
        
        # Try fourth image
        print(f"Attempt {attempt}: Looking for fourth image...")
        if find_and_click_image(imagePath4, "fourth image"):
            clicked_any = True
        else:
            print(f"Attempt {attempt}: Fourth image not found")
        
        if clicked_any:
            print(f"Successful click(s) this attempt. Resetting attempt counter.")
            attempt = 1
        else:
            attempt += 1
        remaining_time = end_time - time.time()
        if remaining_time > 0:
            print(f"Waiting 2 seconds before next attempt... ({remaining_time:.1f}s remaining)")
            time.sleep(2)
    
    if attempt > attemptCount:
        print(f"Reached max attempt count ({attemptCount}). Terminating before timeout.")
    print(f"Timeout reached after {timeout} seconds. Finished clicking attempts.")

# imagePath1 = r"R:\Finance\Revenue Audit\HOTEL\City Ledger Report\Processes\runButton.PNG"
# imagePath2 = r"R:\Finance\Revenue Audit\HOTEL\City Ledger Report\Processes\runButton2.PNG"
# timeout = 30

imagePath1 = sys.argv[1]
imagePath2 = sys.argv[2]
timeout = int(sys.argv[3])
click_images(imagePath1, imagePath2, timeout)