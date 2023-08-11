import ctypes
import ctypes.wintypes
import psutil
import subprocess
import atexit
import signal

# Configuration
APPS = [("teams.exe", "Microsoft Teams"), ("outlook.exe", "Outlook"), ("chrome.exe", "Chrome")] # Configure apps to monitor for
BROWSERS = {'chrome': 'ChromeHTML', 'edge': 'MSEdgeHTM'} # First browser is set for monitored apps
PROTOCOLS = ['http', 'https', '.htm', '.html', '.xhtml', 'ftp', '.svg', '.webp', '.mhtml', '.shtml', '.xml']

DEBUG_MODE = False

# Global variables
current_default_browser = None
hook = None 


# Unhook the Windows event
def unhook_win_event():
    global hook
    if hook:
        ctypes.windll.user32.UnhookWinEvent(hook)

atexit.register(unhook_win_event)  # Register the unhook function to be called on exit


# Get the PID of the active window
def get_active_window_pid():
    active_window_handle = ctypes.windll.user32.GetForegroundWindow()
    active_window_pid = ctypes.wintypes.DWORD()
    ctypes.windll.user32.GetWindowThreadProcessId(active_window_handle, ctypes.byref(active_window_pid))
    return active_window_pid.value


# Execute a given shell command, capturing and handling errors
def execute_command(command):
    try:
        subprocess.run(command, shell=True, capture_output=True, check=True)
    except subprocess.CalledProcessError as e:
        if DEBUG_MODE:
            print(f"Error executing command '{command}': {e.stderr.decode('utf-8')}")
    except Exception as e:
        if DEBUG_MODE:
            print(f"Unexpected error executing command '{command}': {e}")


# Change the default browser to the given one, identified by its name
def change_default_browser(browser):
    global current_default_browser
    browser = browser.lower()
    if browser == current_default_browser:
        return
    prog_id = BROWSERS.get(browser)
    if prog_id:
        for protocol in PROTOCOLS:
            command = 'SetUserFTA {} {}'.format(protocol, prog_id)
            execute_command(command)
        current_default_browser = browser
        if DEBUG_MODE:
            print(f'Changed default browser to {browser}. Program ID: {prog_id}')


# Callback function that handles the foreground window change event
def on_foreground_window_change(hWinEventHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime):
    active_window_pid = get_active_window_pid()
    for process_name, app_name in APPS:
        try:
            process = psutil.Process(active_window_pid)
            if process_name.lower() == process.name().lower():
                if DEBUG_MODE:
                    print(f"{app_name} is currently active")
                change_default_browser('chrome')
                return
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            continue

    if DEBUG_MODE:
        print('Other application is currently active')
    change_default_browser('edge')


# Handle the exit signal to gracefully unhook the event and exit the script
def exit_handler(signal, frame):
    user32.UnhookWinEvent(hook)
    ctypes.windll.user32.PostQuitMessage(0)  # Post a quit message to exit the loop
    print("Exiting...")


# Main function
if __name__ == '__main__':
    # Register the exit handler to handle the SIGINT signal (Ctrl+C)
    signal.signal(signal.SIGINT, exit_handler)

    # Define a Windows function callback type
    WinEventProc = ctypes.WINFUNCTYPE(None, ctypes.wintypes.HANDLE, ctypes.wintypes.DWORD, ctypes.wintypes.HWND, ctypes.wintypes.LONG, ctypes.wintypes.LONG, ctypes.wintypes.DWORD, ctypes.wintypes.DWORD)
    
    # Create the event callback using the defined function type
    event_callback = WinEventProc(on_foreground_window_change)
    
    # Access the user32 library
    user32 = ctypes.windll.user32
    
    # Set the Windows event hook to listen to the foreground window change event
    hook = user32.SetWinEventHook(
        3, # EVENT_SYSTEM_FOREGROUND
        3, # EVENT_SYSTEM_FOREGROUND
        0,
        event_callback,
        0,
        0,
        0 # WINEVENT_OUTOFCONTEXT
    )

    # Create a message object to handle the messages in the queue
    msg = ctypes.wintypes.MSG()

    # Main message loop
    try:
        while ctypes.windll.user32.GetMessageW(ctypes.byref(msg), 0, 0, 0) != 0:
            # Translate and dispatch the messages from the queue
            ctypes.windll.user32.TranslateMessage(ctypes.byref(msg))
            ctypes.windll.user32.DispatchMessageW(ctypes.byref(msg))
    except KeyboardInterrupt:
        # Call the exit handler if a keyboard interrupt is detected
        exit_handler(None, None)
