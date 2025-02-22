import datetime
import os
import subprocess
import sys
import tomllib
from pprint import pprint
from pathlib import Path
import pathlib

import pyautogui as p

p.FAILSAFE = True


# What values do we need set by the user?
# gspro path
# sqg connector path

# Get the directory containing the script
script_path = Path(__file__).resolve()
script_directory = script_path.parent

# Default timeout when waiting for windows to open
WINDOW_TIMEOUT = 10
UI_TIMEOUT = 15

# TODO: add toggle for users to choose between direct exe or launcher
gspro_file_path = Path("C:/GSProV1/Core/GSP/GSPro.exe")
sqg_connector_path = Path("C:/Users", os.getlogin(), "AppData/Local/Programs/Invant Inc/SQG-GSPRO-Connect/SQG-GSPRO-Connect.exe")

# Application window names
open_api_name = "APIv1 Connect"
connector_name = "SQG-GSPRO-Connect"
gspro = "GSPro"

# Images used to navigate the screen
scan_btn = Path(script_directory, "images/scan.png")
connect_btn = Path(script_directory, "images/connect.png")
dropdown_btn = Path(script_directory, "images/populated_dropdown.png")
disconnect_verify = Path(script_directory, "images/disconnect.png")
gsp_connected = Path(script_directory, "images/gsp_connected.png")

# TODO: add custom IP configuration

windows = p.getAllWindows()


def initVars():
  """Load and check the provided config values"""
  with open("pyproject.toml", "rb") as f:
    config = tomllib.load(f).get("config")

  # Don't load if left empty
  if config.gspro_folder != "":
    gspro_file_path_check = Path(config.gspro_folder, "Core/GSP/GSPro.exe").resolve()

  if gspro_file_path_check.exists() and gspro_file_path_check.is_file():
    gspro_file_path = gspro_file_path_check

  # Don't load if left empty
  if config.sqg_connector_path != "":
    sqg_connector_path_check = Path(config.sqg_connector_path)

    if sqg_connector_path_check.exists() and sqg_connector_path_check.is_file():
      sqg_connector_path = sqg_connector_path_check

  # Sanity check defaults also exist
  if not gspro_file_path.exists() or not sqg_connector_path.exists():
    print(f"Installed program check failed.\n"
          + f"\tGSPro: { 'installed' if gspro_file_path.exists() else f'not found at {gspro_file_path}'}\n"
          + f"\tSQG Connector: { 'installed' if sqg_connector_path.exists() else f'not found at {sqg_connector_path}'}"
          )
    sys.exit(1)


def findWindow(name: str):
  """Locate window for given `name` of application.

  Can do it one of two ways.
  1. Get all windows, loop through -- more reliable. 
      `p.getAllWindows()`
  2. Get window directly, sometimes doesn't work, needs name to be exact.
      `p.getWindowsWithTitle(name)[0]`
  """
  windows = p.getAllWindows()
  target_window = None
  for window in windows:
    if name in window.title:
      target_window = window
      return target_window
  return None


def verifyWindow(win_name: str):
  """Verify a given window is open, with a set timeout limit."""
  print(f"Verifying {win_name} window is open...")
  window = findWindow(win_name)

  start_time, end_time = timer(WINDOW_TIMEOUT)
  while window is None:
    p.sleep(1)
    window = findWindow(win_name)
    if datetime.datetime.now() > end_time:
      print(f"{win_name} window could not be verified open.")
      return None
  else:
    print(f"{win_name} window verified open. Continuing...")
    return window


def timer(seconds: int):
  """Create a timer for `seconds` in the future."""
  start_time = datetime.datetime.now()
  end_time = start_time + datetime.timedelta(seconds=seconds)
  return start_time, end_time


def verifyUI(image: Path, grayscale: bool = False, region: tuple[int, int, int, int] = None, click: bool = False):
  """Verify UI element has changed to desired state, optionally clicking"""
  confidence = 0.9
  img_str = str(image)
  start_time, end_time = timer(UI_TIMEOUT)

  elem = None
  while elem is None:
    elem = p.locateOnScreen(image=img_str, grayscale=grayscale, confidence=confidence, region=region)
    if datetime.datetime.now() > end_time:
      print(f"{image.stem} UI could not be found. Exiting...")
      close_apps()
      sys.exit(1)

    p.screenshot(f"verification_screenshots/{image.stem}_verification_screenshot.png", region=elem)

  if click:
    mid = p.center(elem)
    p.click(mid.x, mid.y, duration=1.0)
    # p.click(img_str, duration=1.0)


def verifyUIMulti(image: Path, grayscale: bool = False, region: tuple[int, int, int, int] = None, click: bool = False):
  """Verifies UI for an element with more than one expected to exist. With optional clicking."""
  confidence = 0.9
  img_str = str(image)
  start_time, end_time = timer(UI_TIMEOUT)

  elems = None
  while elems is None:
    elems = p.locateAllOnScreen(image=img_str, grayscale=grayscale, confidence=confidence, region=region)
    if datetime.datetime.now() > end_time:
      print(f"{image.stem} UI could not be found. Exiting...")
      close_apps()
      sys.exit(1)

  if click:
    for pos in elems:
      button_point = p.center(pos)
      p.click(button_point.x, button_point.y, duration=1.0)


def launchGSPro():
  """Launch and verify GSPro and API connector windows open."""
  print("Launching GSPro")
  gspro_pid = subprocess.Popen(gspro_file_path)

  # Wait for it to load
  p.sleep(10)

  # Check if GSPro is open
  gspro_win = verifyWindow(gspro)

  # Check API connector is open
  open_api_win = verifyWindow(open_api_name)

  return gspro_pid


def launchSQGConnector():
  """Launches SQG Connector and verifies it opened."""
  print("Launching SQG Connector")
  connector_process = subprocess.Popen(sqg_connector_path)

  # Check if SQG connector is open
  connector_win = verifyWindow(connector_name)
  return connector_process, connector_win


def automate(sqg_connector_window):
  """
  Automates clicking the SQG Connector and connecting to launch monitor.

  Steps
  1. Click Scan
  2. Wait for dropdown to populate
  3. Click Connect (to LM)
  4. Click Connect (to GSPro)
  5. Verify connected to LM
  6. Verify connected to GSPro
  7. Switch active window to GSPro
  """
  # Get the SQG connector window region to reduce search area for clicks, speeding up process
  rect = (sqg_connector_window.left, sqg_connector_window.top, sqg_connector_window.width, sqg_connector_window.height)
  sqg_connector_window.activate()

  # print(f"connector_window: {connector_window}")
  # print(f"connector_window._rect: {rect}")
  # print("connector_window._rect attributes:")
  # pprint(dir(connector_window._rect))

  # 1. Scan
  verifyUI(scan_btn, region=rect, click=True)
  # scan = p.locateCenterOnScreen(scan_btn, confidence=0.8, region=rect)
  # while scan is None:
  #   scan = p.locateCenterOnScreen(scan_btn, confidence=0.8, region=rect)
  # p.click(scan.x, scan.y, duration=1.0)

  # 2. Wait for dropdown to populate
  # once dropdown is located, LM is found
  # enable grayscale because dropdown can be tinted
  verifyUI(dropdown_btn, grayscale=True, region=rect)
  # dropdown = p.locateOnScreen(dropdown_btn, grayscale=True, confidence=0.8, region=rect)
  # while dropdown is None:
  #   dropdown = p.locateOnScreen(dropdown_btn, grayscale=True, confidence=0.8, region=rect)
  print("Successfully verified LM was found")

  # 3. Click Connect (to LM)
  # 4. Click Connect (to GSPro)
  verifyUIMulti(connect_btn, region=rect, click=True)
  # connect_buttons = p.locateAllOnScreen(connect_btn, confidence=0.8, region=rect)
  # while connect_buttons is None:
  #   connect_buttons = p.locateAllOnScreen(connect_btn, confidence=0.8, region=rect)

  # i = 1
  # for pos in connect_buttons:
  #   print(f"Connect button {i}: {pos}")
  #   button_point = p.center(pos)
  #   p.click(button_point.x, button_point.y, duration=1.0)
  #   print(f"Clicked {i} connect button")
  #   i += 1

  # 5. Verify connected to LM
  verifyUI(disconnect_verify, region=rect)
  # lm_connected = p.locateOnScreen(disconnect_verify, confidence=0.8, region=rect)
  # while lm_connected is None:
  #   lm_connected = p.locateOnScreen(disconnect_verify, confidence=0.8, region=rect)
  print("Successfully verified LM has connected")

  # 6. Verify connected to GSPro
  verifyUI(gsp_connected, region=rect)
  # gsp_connect = p.locateOnScreen(gsp_connected, confidence=0.8, region=rect)
  # while gsp_connect is None:
  #   gsp_connect = p.locateOnScreen(gsp_connected, confidence=0.8, region=rect)
  print("Successfully verified connector has connected to GSPro")

  # 7. Switch active window to GSPro
  gspro_win = findWindow(gspro)

  if gspro_win:
    gspro_win.activate()
    print(f"Switched to {gspro}")
  else:
    print(f"{gspro} not found")


def close_apps():
  # attempt to find each window and close them
  apps = [gspro, open_api_name, connector_name]
  windows = p.getAllWindows()

  for window in windows:
    for app in apps:
      if app in window.title:
        window.close()


def main():
  gspro_process = launchGSPro()
  connector_process, connector_window = launchSQGConnector()
  automate(connector_window)

  print("Watching for GSPro to exit, will close connector and API windows.")
  if gspro_process.wait() is not None:
    # close_apps()
    connector_process.kill()
    print("Closed other apps. Goodbye.")


if __name__ == "__main__":
  main()
