import pyautogui
import time
import keyboard
import random
import win32api, win32con


snow_link_loc = pyautogui.locateOnScreen(r'C:\Users\mathew.roberts\Desktop\Incidents Icon.PNG', region = (670, 60, 300, 50))
no_records_loc = pyautogui.locateOnScreen(r'C:\Users\mathew.roberts\Desktop\No Records Icon.PNG', region = (900, 300, 500, 100))
snow_link_point = pyautogui.center(snow_link_loc)
snow_x = snow_link_point[0]
snow_y = snow_link_point[1]

def smooth_mouse_move(x, y, t):
    
    pyautogui.moveTo(x, y, t)

def left_click(x, y):
    win32api.SetCursorPos((x, y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)

def right_click(x, y):
    win32api.SetCursorPos((x, y))
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0)
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0)

def launch_snow():
    smooth_mouse_move(snow_link_point[0], snow_link_point[1], 0.01)
    left_click(snow_x, snow_y)

if snow_link_loc != None:
    launch_snow()
    time.sleep(3)
    if no_records_loc != None:
        print('There are currently tickets to process.')
    else:
        print('There are no tickets in the queue.') 