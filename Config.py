import os
from posixpath import split
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

currentdir = os.path.dirname(__file__)
groups_list = {}
location_drive_dict = {}
file_path_edge_seal = 'C:\My Drive\Edge_Seal_Inspection_Image\EdgeSealInspection_GlassToBead.png'
file_path_cover_glass = 'C:\My Drive\CoverGlassRobot_Alarm_Image File\CoverGlassRobot_Alarm.png'
message = 'Fin 1 Update Data'
Master_Folder = 'C:\My Drive'
drive_path = Master_Folder

# SQL Source Code
sql_sourcecode_list ={}
with open(os.path.join(os.path.dirname(__file__), "Sql_Source_Code"), "r") as source:
    sql_sourcecode_list = {
        _.split(":")[0].strip().split("_")[0].strip('SQL'): _.split(":")[1].strip() 
        for _ in source.read().split(";")
        if len(_) > 0
        # Lọc dữ liệu trong dictionary
    }

# Receiver And Group
receiver_group_name = []
with open(os.path.join(os.path.dirname(__file__), "Receiver_Group_Name"), "r", encoding='utf8') as source:
    receiver_group_name = [_.strip() for _ in source.readlines()]

# Data Storage Location
location_drive_dict = []
with open(os.path.join(os.path.dirname(__file__), "DataLocationPath"), "r", encoding='utf-8') as datapath:
    location_drive_dict = {
        _.split("=")[0].strip():_.split("=")[1].strip()
            for _ in datapath.read().strip().split(";")
            if len(_) > 0
    }

# Get DataName From SQL_Source_Code
dataname_list = []
with open(os.path.join(os.path.dirname(__file__),"SQL_Source_Code"), "r", encoding='utf-8') as source:
    dataname_list = [
        _.split(":")[0].strip().split("_")[0].strip('SQL')
        for _ in source.read().split(";")
        if len(_) > 0]

# Terminate The Process After Plotting Chart
def Terminate_Process (name):
    os.system("taskkill /f /im {} ".format(name))
    # The OS module in Python provides functions for interacting with the operating system.
    # f : Specifies that process(es) be forcefully terminated.
    # /im (ImageName ) : Specifies the image name of the process to be terminated.

# Initalizing Chrome And Access To Whatsapp
def initalizing_chrome():
    url = "https://web.whatsapp.com/"
    chromedirectory = "user-data-dir=C:\\Users\\fs120806\AppData\Local\Google\Chrome\\User Data"
    options = webdriver.ChromeOptions()
    options.add_argument(chromedirectory)
    initalizing_chrome.driver = webdriver.Chrome(executable_path="C:\Chrome Driver\chromedriver.exe", options=options)
    initalizing_chrome.driver.maximize_window()
    print("Start Sending Message")
    initalizing_chrome.driver.get(url)

# Get All DataName And Hour In Sql_Source_Code
with open(os.path.join(os.path.dirname(__file__), "Sql_Source_Code"), "r") as source:
    dataname_hour = {
        _.split("SQL_")[0].strip():_.split("SQL_")[1].strip().split(":")[0]
        for _ in source.read().split(";")
        if len(_) > 0
    }
