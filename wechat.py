import time

import pyautogui
import win32api
import win32clipboard
import win32con
import win32gui


def find_window(name):
    window = win32gui.FindWindow(None, name)
    if window != 0:
        win32gui.ShowWindow(window, win32con.SW_SHOWMINIMIZED)
        win32gui.ShowWindow(window, win32con.SW_SHOWNORMAL)
        win32gui.ShowWindow(window, win32con.SW_SHOW)
        win32gui.SetWindowPos(window, win32con.HWND_TOPMOST, 0, 0, 300, 500, win32con.SWP_SHOWWINDOW)
        win32gui.SetForegroundWindow(window)
        time.sleep(0.25)
        return True
    else:
        print(f"Error: {name} not found")
        return False


def set_clipboard_text(text):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
    win32clipboard.CloseClipboard()
    time.sleep(0.1)
    win32api.keybd_event(17, 0, 0, 0)
    win32api.keybd_event(86, 0, 0, 0)
    win32api.keybd_event(86, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)


def click_send():
    win32api.keybd_event(18, 0, 0, 0)
    win32api.keybd_event(83, 0, 0, 0)
    win32api.keybd_event(18, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(83, 0, win32con.KEYEVENTF_KEYUP, 0)


def is_wechat():
    pyautogui.click(25, 110)
    return win32gui.GetPixel(win32gui.GetDC(win32gui.GetActiveWindow()), 25, 110) == 6340871


# send text to wechat
def send_wechat(wechat, text):
    try:
        # find wechat window
        if find_window('WeChat'):
            time.sleep(0.2)
        else:
            return
        if not is_wechat():
            print("Error: this is not wechat")
            return

        # navigate to search box
        pyautogui.moveTo(143, 39)
        try:
            pyautogui.click()
        except Exception as e:
            print(f'error: {e}')
            return

        # search wechat name
        set_clipboard_text(wechat)
        time.sleep(0.6)

        # navigate to text box
        pyautogui.moveTo(155, 120)
        try:
            pyautogui.click()
        except Exception as e:
            print(f'error: {e}')
            return

        # paste and send
        set_clipboard_text(text)
        click_send()

        print(f'Successfully send "{text}" to "{wechat}"')
    except Exception as e:
        print(f'Error: {e}')
