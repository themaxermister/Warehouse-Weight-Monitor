## Warehouse Weight Monitor v1.02

import os
import sys
import serial
import _thread
import time
import webbrowser

import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

import selenium
from selenium.webdriver.remote.command import Command
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

class SerialPort:
    def __init__(self):
        self.comportName = ""
        self.baud = 0
        self.timeout = None
        self.ReceiveCallback = None
        self.isopen = False
        self.receivedMessage = None
        self.serialport = serial.Serial()

    def __del__(self):
        try:
            if self.serialport.is_open():
                self.serialport.close()
        except:
            print("Destructor error closing COM port: ", sys.exc_info()[0] )

    def RegisterReceiveCallback(self,aReceiveCallback):
        self.ReceiveCallback = aReceiveCallback
        try:
            _thread.start_new_thread(self.SerialReadlineThread, ())
        except:
            print("Error starting Read thread: ", sys.exc_info()[0])

    def SerialReadlineThread(self):
        while True:
            try:
                if self.isopen:
                    self.receivedMessage = self.serialport.readline()
                    if self.receivedMessage != "":
                        self.ReceiveCallback(self.receivedMessage)
            except:
                print("Error reading COM port: ", sys.exc_info()[0])
                weightbox.config(text="RECONNECT & RESTART")

    def IsOpen(self):
        return self.isopen

    def Open(self,portname,baudrate):
        if not self.isopen:
            # serialPort = 'portname', baudrate, bytesize = 8, parity = 'N', stopbits = 1, timeout = None, xonxoff = 0, rtscts = 0)
            self.serialport.port = portname
            self.serialport.baudrate = baudrate
            try:
                self.serialport.open()
                self.isopen = True
                return True
            except:
                print("Error opening COM port: ", sys.exc_info()[0])
                self.isopen = False
                return False


    def Close(self):
        if self.isopen:
            try:
                self.serialport.close()
                self.isopen = False
            except:
                print("Close error closing COM port: ", sys.exc_info()[0])

    def Send(self,message):
        if self.isopen:
            try:
                # Ensure that the end of the message has both \r and \n, not just one or the other
                newmessage = message.strip()
                newmessage += '\r\n'
                self.serialport.write(newmessage.encode('utf-8'))
            except:
                print("Error sending message: ", sys.exc_info()[0] )
            else:
                return True
        else:
            return False

# GLOBALS
## Test Values
## website = "https://www.google.com/"
## x_path =   "//*[@id='tsf']/div[2]/div[1]/div[1]/div/div[2]/input"

serialPort = SerialPort()
website =   ## TARGET WEBSITE
x_path =    ## TARGET WEB ELEMENT
com_port = "COM3"

# COOKIES
path = os.getcwd()
chrome_cookie_argument = f"{path}/cookies"    ## COOKIES FOLDER

capabilities = {
    'chromeOptions': {
        'useAutomationExtension': False,
        'args': ['--start-maximized',
                 '--disable-infobars',
                 '--disable-notifications',
                 '--disable-popup-blocking',
                 f"user-data-dir={chrome_cookie_argument}"
                 ]
    }
}

driver = webdriver.Chrome(
   f'{path}/chromedriverUK.exe', desired_capabilities=capabilities)   ## CHROME DRIVER FILE
driver.get(website)

# TK ROOT WINDOW
root = tk.Tk()
root.title("Warehouse Weight Monitor v1.01" )
root.configure(bg = '#ffffff')
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
window_width = screen_width/2
window_height = screen_width/3
window_position_x = screen_width/2 - window_width/2
window_position_y = screen_height/2 - window_height/2
root.geometry('%dx%d+%d+%d' % (window_width, window_height, window_position_x, window_position_y))

# SEND DATA TO WEB UI
def SendData(weight):
    global driver
    if serialPort.IsOpen():
        try:
            _ = driver.window_handles

        except selenium.common.exceptions.WebDriverException as e:
            driver = webdriver.Chrome(
                f'{path}/chromedriverUK.exe', desired_capabilities=capabilities)
            driver.get(website)
            weight = "RETRY"
            weightbox.config(text = "ลองใหม่")  # Retry
            return

        try:
            weight = weight.replace(" kg", '')
            test = float(weight)
            xpath = x_path  ## CHANGE
            form_element = driver.find_element_by_xpath(xpath)
            form_element.send_keys(weight)
            form_element.send_keys(Keys.ENTER)
        
        except ValueError:
            weightbox.config(text = "ERROR")

        except selenium.common.exceptions.NoSuchElementException:
            weightbox.config(text = "ผิดหน้า") # Wrong Page
        
        except selenium.common.exceptions.InvalidElementStateException:
            weightbox.config(text = "เลือกคำสั่งซื้อและสินค้าก่อน") # Select an order first
    else:
        print("Invalid")
        return

# # Close programs
def Close():    
    try:
        driver.close()
    except selenium.common.exceptions.WebDriverException as e:
        print("Browser not running")

    exit()

# CHECK STRING FORMAT
def check_format(output):
    str_form = re.search("1,ST,     (.+?),       0.000,         0,kg", output)
    if str_form:
        found = str_form.group(1)
        return found

    str_form2 = re.search('GROSS(.+?)kg', output)
    if str_form2:
        found = str_form2.group(1)
        return found

# SERIAL DATA CALLBACK
def OnReceiveSerialData(message):
    if not serialPort.IsOpen():
        weightbox.config(text = "RECONNECT & RESTART")
    else:
        str_message = message.decode("utf-8")
        empty = str_message.isspace()
        if (not empty):
            print(str_message + "\n")
            weight = (check_format(str_message)).replace(" ", "")
            weightbox.config(text = weight + " kg")
            SendData(weightbox.cget('text'))
            return weight

# # Register the callback above with the serial port object
serialPort.RegisterReceiveCallback(OnReceiveSerialData)

# # Run main loop once every 200 ms
def sdterm_main():
    root.after(200, sdterm_main)

def OpenCommand():
    global driver

    comport = com_port
    baudrate = 9600
    serialPort.Open(comport,baudrate)

    if not serialPort.IsOpen():
        weightbox.config(text = "กรุณาเช็คการเชื่อมต่อสาย") # Port Close
        OpenCommand()

    else:
        weightbox.config(text = "ชั่งน้ำหนัก")  # Weight Here

# BUTTON FUNCTIONS
# UI DESIGN
# # Labels
# ## Label box containing weight
weightbox = tk.Label(root, fg = '#fc5603', borderwidth=6, relief="ridge")
weightbox.config(bg='white', font=('Helvetica', 40))
weightbox.place(relx=0.5, rely=0.35, anchor=CENTER)

# # Buttons
# ## Close Program
button_close = Button(root,text="ออก",width=20,command=Close)  # Exit
button_close.config(font="bold")
button_close.place(relx=0.5, rely=0.6, anchor=CENTER)

# MAIN FUNCTION
OpenCommand()
root.after(200, sdterm_main)
root.mainloop()
