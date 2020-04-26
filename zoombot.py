#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import win32gui
import win32api
import win32con
import subprocess
import sys
import pynput
import time
import schedule

ZOOM_PATH = "C:\\Users\\coolguy\\AppData\\Roaming\\Zoom\\bin\\Zoom.exe"

DEFAULT_USERNAME = "Tarık"
DEFAULT_PASSWORD = "123456"
MEETINGS = {
    "kimya": "xxx-xxx-xxx",
    "biyoloji": "xxx-xxx-xxx",
    "fizik": "xxx-xxx-xxx",
    "edebiyat": "xxx-xxx-xxx",
    "mat1:": "xxx-xxx-xxx",
    "mat2": "xxx-xxx-xxx"
}

def move_zoom_window(xpos=0, ypos=0, width=300, height=400, appname="Zoom"):
    def enumHandler(hwnd, lParam):
        if win32gui.IsWindowVisible(hwnd):
            if appname in win32gui.GetWindowText(hwnd):
                win32gui.MoveWindow(hwnd, xpos, ypos, width, height, True)
    win32gui.EnumWindows(enumHandler, None)
    if (appname == "Zoom"):
        move_zoom_window(xpos=xpos, ypos=ypos, width=width, height=height, appname="Zoom Cloud Meeting")
    if (appname == "Zoom Cloud Meeting"):
        move_zoom_window(xpos=xpos, ypos=ypos, width=width, height=height, appname="Enter meeting password")

def start_zoom(path):
    subprocess.Popen([ZOOM_PATH], shell=False, stdin=None, stdout=None, stderr=None, close_fds=True)

def close_zoom():
    subprocess.Popen(["taskkill", "/f", "/im", "Zoom.exe"], shell=True, stdin=None, stdout=None, stderr=None, close_fds=True)
    time.sleep(1)

def click(x, y):
    controller = pynput.mouse.Controller()
    controller.position = (x, y)
    controller.click(pynput.mouse.Button.left, 1)

def write(str):
    controller = pynput.keyboard.Controller()
    controller.press(pynput.keyboard.Key.ctrl_l.value)
    controller.press("a")
    controller.release("a")
    controller.release(pynput.keyboard.Key.ctrl_l.value)
    controller.press(pynput.keyboard.Key.delete.value)
    controller.type(str)

def join_meeting(meeting_id, name, password, meeting_min=60, width=300, height=400):
    print("[+] Closing old Zoom")
    close_zoom()
    time.sleep(3)
    print("[+] Opening new Zoom")
    start_zoom(ZOOM_PATH)
    time.sleep(3)   
    print("[+] Resizing zoom windows")
    move_zoom_window()
    
    #Click on join meeting button
    time.sleep(3)
    print("[+] Clicking on join meeting button")
    join_button_x = int(width / 2)
    join_button_y = int(height / 2)
    click(join_button_x, join_button_y)
    
    time.sleep(3)
    print("[+] Resizing zoom windows")
    move_zoom_window()
    
    # Set meeting id
    meeting_id_x = int(width / 2)
    meeting_id_y = int((height * 1.75) / 5)
    click(meeting_id_x, meeting_id_y)
    print("[+] Writing meeting id: {}".format(meeting_id))
    write(meeting_id)
   
    # Set user name
    username_x = int(width / 2)
    username_y = int((height * 2.5) / 5)
    click(username_x, username_y)
    print("[+] Writing username: {}".format(name))
    write(name)
    
    # Click join
    join_x = int(width / 2)
    join_y = int((height * 4.25) / 5)
    click(join_x, join_y)
    print("[+] Joining to the meeting")
    time.sleep(3)
    move_zoom_window()
    
    # Write password
    password_x = int(width / 2)
    password_y = int((height * 1.75) / 5)
    click(password_x, password_y)
    print("[+] Writing meeting password: {}".format(password))
    write(password)
    
    # Click join
    join_x = int(width / 2)
    join_y = int((height * 4.25) / 5)
    click(join_x, join_y)
    print("[+] Joining to the meeting")
    time.sleep(3)
    move_zoom_window()
    
    timer = 0
    while (timer != meeting_min):
        move_zoom_window()
        time.sleep(15)
        move_zoom_window()
        time.sleep(15)
        move_zoom_window()
        time.sleep(15)
        move_zoom_window()
        time.sleep(15)
        timer = timer + 1
    
    print ("[+] Meeting has ended!")
    close_zoom()
    
def enter_class(lesson):
    meeting_id = MEETINGS[lesson]
    try:
        join_meeting(meeting_id, DEFAULT_USERNAME, DEFAULT_PASSWORD, meeting_min=39)
    except e:
        print(e)
        enter_class(lesson)

def set_schedules():
    print("[+] Setting schedules")
    schedule.every().monday.at("09:30").do(enter_class, lesson="kimya")
    schedule.every().monday.at("10:20").do(enter_class, lesson="kimya")
    schedule.every().monday.at("11:10").do(enter_class, lesson="mat1")
    schedule.every().monday.at("12:00").do(enter_class, lesson="biyoloji")
    schedule.every().monday.at("12:50").do(enter_class, lesson="fizik")
    
    schedule.every().tuesday.at("09:30").do(enter_class, lesson="mat2")
    schedule.every().tuesday.at("10:20").do(enter_class, lesson="mat2")
    schedule.every().tuesday.at("11:10").do(enter_class, lesson="fizik")
    schedule.every().tuesday.at("12:00").do(enter_class, lesson="fizik")
    schedule.every().tuesday.at("12:50").do(enter_class, lesson="fizik")
    
    schedule.every().wednesday.at("09:30").do(enter_class, lesson="mat2")
    schedule.every().wednesday.at("10:20").do(enter_class, lesson="mat2")
    schedule.every().wednesday.at("11:10").do(enter_class, lesson="edebiyat")
    schedule.every().wednesday.at("12:00").do(enter_class, lesson="mat1")
    schedule.every().wednesday.at("12:50").do(enter_class, lesson="biyoloji")
    
    schedule.every().thursday.at("09:30").do(enter_class, lesson="mat2")
    schedule.every().thursday.at("10:20").do(enter_class, lesson="mat2")
    schedule.every().thursday.at("11:10").do(enter_class, lesson="kimya")
    schedule.every().thursday.at("12:00").do(enter_class, lesson="mat1")
    schedule.every().thursday.at("12:50").do(enter_class, lesson="biyoloji")
    
    schedule.every().friday.at("09:30").do(enter_class, lesson="edebiyat")
    schedule.every().friday.at("10:20").do(enter_class, lesson="kimya")
    schedule.every().friday.at("11:10").do(enter_class, lesson="kimya")
    schedule.every().friday.at("12:00").do(enter_class, lesson="biyoloji")
    schedule.every().friday.at("12:50").do(enter_class, lesson="mat2")
    
    print("[+] Waiting for class")

if __name__ == "__main__":
    set_schedules()
    while True:
        schedule.run_pending()
        time.sleep(1)