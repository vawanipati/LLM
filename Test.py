"""
cycle_windows_tab_64.py
Windows 10/11 (64-bit) – No admin required.
Cycles visible windows to the front every 5 minutes using SetForegroundWindow.
If focus fails, simulates Alt+Tab.

Install first:
    pip install pywin32 pyautogui
"""

import time
import win32gui
import win32con
import pyautogui


INTERVAL = 300  # 5 minutes (change to 5 for quick test)


def get_visible_windows():
    """Return list of (hwnd, title) for visible top-level windows with titles."""
    windows = []

    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title.strip():
                windows.append((hwnd, title))
        return True

    win32gui.EnumWindows(callback, None)
    return windows


def bring_to_front(hwnd, title):
    """Try to bring window to front; if not allowed, simulate Alt+Tab."""
    try:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
        print(f"✔️  Focused: {title}")
    except Exception:
        print(f"⚠️  Could not focus {title}. Using Alt+Tab fallback.")
        pyautogui.keyDown('alt')
        pyautogui.press('tab')
        pyautogui.keyUp('alt')


def main():
    print("=== Window Cycler for Windows 64-bit (no admin needed) ===")
    print(f"Interval: {INTERVAL} seconds\n")

    try:
        while True:
            windows = get_visible_windows()
            if not windows:
                print("No visible windows found. Retrying in 60 seconds...")
                time.sleep(60)
                continue

            print(f"Found {len(windows)} windows.")
            for hwnd, title in windows:
                bring_to_front(hwnd, title)
                time.sleep(INTERVAL)
    except KeyboardInterrupt:
        print("\nStopped by user.")


if __name__ == "__main__":
    main()
