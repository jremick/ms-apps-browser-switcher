# Microsoft Apps Default Browser Switcher
I use a different browser for work vs personal and couldn't find an easy way for URLs from Microsoft Office and Teams to open in a specific browser, while all other apps' URLs open in a different browser. 

This Python script monitors specified applications' active windows (Microsoft Office and Teams) and changes the default browser based on which application is in focus for > 5 seconds. Google Chrome is used for Outlook and Teams, Edge is used for everything else.

### Features

- Monitors specified applications' active windows.
- Switches the default browser based on the focused application.
- Reduces CPU usage by adjusting the check interval.

### Requirements

1. Python 3.x
2. Required Python Libraries:
   - `psutil`
   - `win32gui`
   - `win32process`
  
Written for use on Windows 11, but should also work on Windows 10.
  
### How to Use

1. Set up the environment:
    ```
    pip install psutil pywin32
    ```

2. Modify the list of applications to monitor by adjusting the `apps` list in the script.
3. If `DEBUG_MODE` is set to `True`, the script will print debug messages. Set to `False` to disable.
4. Run the script:
    ```
    python script_name.py
    ```

### Script Behavior

1. Monitors the active window of specified applications: Microsoft Teams, Outlook, and Chrome by default.
2. If one of the specified applications is active for more than 5 seconds, the default browser is set to Chrome. (Chrome must be in the monitored list so that while _in_ Chrome, the default browser isn't reset to Edge.) 
3. If any other application is active for more than 5 seconds, the default browser is set to Edge.
4. The script uses `SetUserFTA` command to change the default browser. Ensure you have the necessary permissions to change file type associations.

### Note

The `SetUserFTA` command and its functionality are not included in the script. Ensure you have this utility available and adjust the paths as needed.

## Compiling into a Standalone Executable for Windows 11

To turn the script into a standalone Windows executable, you can use `PyInstaller`. This will allow you to run the application without requiring a Python environment.

### Steps:

1. Install `PyInstaller`:
    ```
    pip install pyinstaller
    ```

2. Navigate to the directory containing your script and compile it:
    ```
    pyinstaller --onefile script_name.py
    ```

   - The `--onefile` flag ensures the output is a single executable file.
   
3. Once the compilation process is complete, you'll find the standalone executable in the `dist` directory within your current directory.

