import time
import win32gui
import win32process
import psutil
import subprocess
import os

# Set the debug mode to false (by default)
DEBUG_MODE = True

# Get the current process
process = psutil.Process(os.getpid())

# Set the process to idle priority class
process.nice(psutil.IDLE_PRIORITY_CLASS)

# Define the applications to monitor
apps = [("teams.exe", "Microsoft Teams"), ("outlook.exe", "Outlook"), ("chrome.exe", "Chrome")]

# Initialize dictionaries to keep track of which applications were active first
app_active_first_check = {app[1]: 0 for app in apps}
current_default_browser = None

# Initialize a set instead of a list for faster lookups
app_process_names = {app[0].lower() for app in apps}

# Function to check if a process is running and get the PID of the active window
def get_running_processes_and_active_window_pid():
    proc_info = [(proc.info['name'].lower(), proc.info['pid']) for proc in psutil.process_iter(['name', 'pid'])]
    running_processes = [(name, pid) for name, pid in proc_info if name in app_process_names]
    # Get the active window PID only if the current process matches one of the monitored processes
    active_window_handle = win32gui.GetForegroundWindow()
    active_window_pid = win32process.GetWindowThreadProcessId(active_window_handle)[1] if any(pid == win32process.GetWindowThreadProcessId(active_window_handle)[1] for name, pid in running_processes) else None
    return running_processes, active_window_pid

# Function to execute a command and return its output
def execute_command(command):
    output = subprocess.run(command, shell=True, capture_output=True)
    if output.stderr:
        if DEBUG_MODE:
            print('Error:', output.stderr.decode('utf-8'))

# Function to change the default browser
def change_default_browser(browser, current_default_browser):
    browser = browser.lower()
    if browser == current_default_browser:
        return current_default_browser  # return current_default_browser if it is already set correctly

    prog_id = {'chrome': 'ChromeHTML', 'edge': 'MSEdgeHTML'}.get(browser)
    if prog_id:
        for protocol in ['http', 'https', '.htm', '.html']:
            command = 'SetUserFTA {} {}'.format(protocol, prog_id)
            execute_command(command)

        current_default_browser = browser  # update the current default_browser
    elif current_default_browser is None:
        if DEBUG_MODE:
            print('Unsupported browser:', browser)

    return current_default_browser  # return the updated current_default_browser


active_window_pid_prev = None
active_window_check_interval = 1.5  # initial sleep time

if __name__ == '__main__':
    while True:
        app_running = False
        running_processes, active_window_pid = get_running_processes_and_active_window_pid()

        for process_name, app_name in apps:
            app_processes = [(name, pid) for name, pid in running_processes if name == process_name.lower()]

            if app_processes:
                app_running = True
                if any(pid == active_window_pid for name, pid in app_processes):
                    if time.time() - app_active_first_check[app_name] > 5:  # app has been active for more than 5 seconds
                        if DEBUG_MODE:
                            print('{} is currently active'.format(app_name))
                        app_active_first_check[app_name] = time.time()
                        current_default_browser = change_default_browser('chrome', current_default_browser)
                    break
        else:  # if the loop completes without breaking (i.e., no app in the list is in focus)
            if time.time() - max(app_active_first_check.values()) > 5:  # no app has been active for more than 5 seconds
                if DEBUG_MODE:
                    print('Other application is currently active')
                app_active_first_check = {app[1]: 0 for app in apps}
                current_default_browser = change_default_browser('edge', current_default_browser)

        # If the active window PID hasn't changed, increment the sleep time to reduce CPU usage
        if active_window_pid == active_window_pid_prev:
            active_window_check_interval = min(active_window_check_interval + 0.1, 3)  # max sleep time of 2 seconds
        else:
            active_window_check_interval = 1.5  # reset sleep time if the active window changes

        active_window_pid_prev = active_window_pid
        time.sleep(active_window_check_interval)  # check for active window changes every second
