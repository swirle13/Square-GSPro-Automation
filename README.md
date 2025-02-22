# Square-GSPro-Automation

Script to automate the annoying steps of using Square connector with GSPro.

This only works for windows as GSPro is windows-only.

**NOTE:** For my first ~~guinnea pigs~~ users, you may need to screenshot your own SQG connector dropdown box and replace [populated_dropdown.png](images/populated_dropdown.png) with your screenshot. This may be due to the Square ID being different per device and may affect the program's ability to successfully find/validate the LM has been found when scanning with bluetooth.

- [Square-GSPro-Automation](#square-gspro-automation)
  - [Brief instructions for those familiar with Python](#brief-instructions-for-those-familiar-with-python)
  - [Instructions for those not familiar with Python](#instructions-for-those-not-familiar-with-python)
  - [Shortcuts](#shortcuts)
    - [Desktop](#desktop)
    - [Start Menu](#start-menu)
  - [Configuration](#configuration)
    - [Finding your application's path](#finding-your-applications-path)

## Brief instructions for those familiar with Python

1. Install or have Python 3.11 or greater installed
2. Clone this repo locally or download the zip, save it anywhere you like
3. Install the project dependencies via `pip` or running the included .bat file, `install_dependencies.bat`
4. If you changed the install location for GSPro or Square Golf, update `pyproject.toml`'s `config` to match your local configuration. See [Configuration](README.md#configuration) for more information.
5. You can now run `main.py` to run GSPro whenever you would like to have the setup/teardown automated. There are a few ways you can do this. See [Shortcuts](README.md#shortcuts) for more information.

## Instructions for those not familiar with Python

1. Install Python 3.11 from the windows store - [link](https://apps.microsoft.com/detail/9nrwmjp3717k?hl=en-US&gl=US)
2. Click the big green "Code" button at the top of this page, then click "Download ZIP"
3. Open your downloaded zip file and extract it wherever you prefer. I recommend putting it in your "My Documents" folder, located at `C:\Users\%username%\Documents`
4. Double click `install_dependencies.bat`
5. If you changed the install location for GSPro or Square Golf, update `pyproject.toml`'s `config` to match your local configuration. See [Configuration](README.md#configuration) for more information.
6. You can now run `main.py` to run GSPro whenever you would like to have the setup/teardown automated. There are a few ways you can do this. See [Shortcuts](README.md#shortcuts) for more information.

## Shortcuts

Create a shortcut to `main.py` and place the shortcut wherever is most convenient for you.

### Desktop

If you want the shortcut on your desktop, simply right click `main.py` and click "Send to" > "Desktop (create shortcut)"

### Start Menu

If you want the shortcut in your start menu, it takes a couple more steps.

1. Right click `main.py` > "Create shortcut"
2. Copy/paste the newly created shortcut to `C:\Users\%username%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs`
3. [Optional] Now open Start Menu > All apps and locate the shortcut you placed. Right-click on it and select Pin to Start.

If you're using Sunlight, like I am, you can set your application command to point to `main.py` instead of `GSPLauncher.exe` or `GSPro.exe`

## Configuration

There are a few values that you can set through the configuration file, `pyproject.toml` under the `gspro` section.

Name | Default | Description
--- | --- | ---
`gspro_folder` | "" (empty) | Empty allows the code to set the default. Otherwise, value should be the **folder** of GSPro. This allows execution of `GSPro.exe` directly instead of `GSPLauncher.exe`, which is quicker and one less window. Requires double backslashes.<br>Example value: `"C:\\User\\Jim\\GSProV1"`
`sqg_connector_path` | "" (empty) | Empty allows the code to set the default. Value should be full path to `SQG-GSPRO-Connect.exe`. Requires double backslashes.<br>Example value: `"C:\\User\\Jim\\SquareGolf\\SQG-GSPRO-Connect.exe"`
`window_timeout` | 10 | The length of time in seconds to set the length of time to time out when searching for an application window. Increase this value if you run into issues with the script being unable to find the window in time. If your computer finds the windows quickly, there's no benefit to lowering this value as it will always continue the moment a window is found. Cannot be set to a negative or zero value
`ui_timeout` | 15 | Time in seconds for UI searches to stop searching. This governs the time allowed when scanning for SQG Connector buttons and dots to change. This includes waiting for a Square to be located, waiting for the LM to connect once "connect" has been clicked, etc.<br>**NOTE:** Square takes about 20s to boot up, so please make sure your square is turned on before running this script, otherwise, increase this number to around 30.

### Finding your application's path

If you need to double check where GSPro or SQG-GSPRO-Connect is installed, you can do the following to get the values.

1. Press the Win key or open up start menu search
2. Type in the program name
3. Right click the application and click "Open file location"
4. If the icon is a shortcut (little arrow in bottom left of icon), right click and select "open file location"
5. If it took you to the file, you can skip to step 7
6. If it took you to the same location you were in, you need to right click the shortcut and select "Properties" then "Shortcut" tab. Copy the path in "Start in" and paste it into the top bar in File Explorer
7. You can now Shift + right click the .exe and select "Copy as path"
8. You can now paste the value anywhere you need.
