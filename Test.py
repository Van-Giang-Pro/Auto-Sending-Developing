import os
from posixpath import split
import pyodbc
import pandas as pd
from datetime import date, datetime, timedelta
import sys
import matplotlib.pyplot as plt
import shutil
import seaborn as sns
import numpy as np
import subprocess
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException

'''
string = 'Giang'
s = f'{string} va Hanh Dung'
print(s)

path = '/Giang/Tien/Dung/Như'
s1 = os.path.dirname(path)
s2 = os.path.basename(path)
print(s1)
print(s2)

a = os.path.dirname(__file__) # trả về đường dẫn chương trình python
b = os.path.basename(__file__) # trả về tên chương trình python
print(a)
print(b)
name = input("Nhap : ")
print(type(name))

def inra(a, b):
    c = a+ ' ' +b
    print(c)
inra('Giang', 'Dung')

str = 'Giang'
print(str.strip('g'))

html_file = open('HTML Creation','w')
html_file.write('Nguyen Van Giang')

with open(os.path.join(os.path.dirname(__file__), "SQL_Source_Code"), "r") as source:
    y = source.read().split(";")
    print(y)
    print(len(y))
    print(len(source.read().split(";")))
    print(type(y))
 
test = 'Nguyen Van Giang :Dang Hoang Hanh Dung'
test1 = test.split(":")[0].strip()
test2 = test.split(":")[1].strip()
print(test1)
print(test2)
print(len(test.split(":")))

g = ['A','B','C','D']
print(len(g))

with open(os.path.join(os.path.dirname(__file__), "SQL_SourceCode"), "r") as source:
    sql_sourcecode_list = {
        _.split(":")[0].strip(): _.split(":")[1].strip()
        for _ in source.read().split(";")
        if len(_) > 0
    }
print(sql_sourcecode_list)
print(type(sql_sourcecode_list))
a = 'Giang'
sql = {}
sql = [a+'Dung']
print(type(sql))

datasource = sql_sourcecode_list['EdgeSealInspectionSQL']\
    .replace("__SelectedSite__", 'SelectedSite')\
    .replace("__StartTime__", 'StartTime')\
    .replace("__EndTime__", 'EndTime')
# Cách truy cập đến phần tử key trong Dict
# \ dùng để cho biết lệnh còn tiếp tục ở dòng dưới cùng
print(datasource) 
print(type(datasource))

def DataCollection(DataName, SelectedSite, StartTime, EndTime):
    server = f"Driver=SQL Server;Server={SelectedSite}MESSQLODS;Database=ODS;Trusted_Connection=Yes"
    datasource = sql_sourcecode_list[DataName+'SQL']\
                .replace("__SelectedSite__", SelectedSite)\
                .replace("__StartTime__", StartTime)\
                .replace("__EndTime__", EndTime)
    connection = pyodbc.connect(server)
    connection.cursor()
    print(datasource)
    DataCollected = pd.read_sql_query(datasource, connection, index_col=None)
    print(DataCollected.head())
    return DataCollected
DataCollection('EdgeSealInspection', 'DMT2', '2021-08-09 02:00:00', '2021-08-09 04:00:00')

# Using current time
ini_time_for_now = datetime.now()
  
# printing initial_date
print ("initial_date", str(ini_time_for_now))
  
# Calculating future dates
# for two years
future_date_after_2yrs = ini_time_for_now + \
                        timedelta(days = 730)
  
future_date_after_2days = ini_time_for_now + \
                         timedelta(days = 2)
                         
# printing calculated future_dates
print('future_date_after_2yrs:', str(future_date_after_2yrs))
print('future_date_after_2days:', str(future_date_after_2days))

print(timedelta(days=2))
t1 = timedelta(weeks = 1, days = 6, hours = 1, seconds = 33)
# timedelta là khoảng thời gian, 1 tuần cộng với 6 ngày cộng 1 giờ và thêm 33 giây
print(t1)
print(datetime.today())
print(type(datetime.today()))
Initial_DateTime = datetime.today().replace(minute=0, second=0)
print(Initial_DateTime)
print(datetime.today())
early_morning = datetime.combine(date.today(),datetime.min.time())
# datetime.combine là kết hợp giữa thời gian và giờ, min là lấy giá trị nhỏ nhất của thời gian
print(early_morning)
print(datetime.min.time())
print(datetime.date(datetime.now()))

with open('Receiving Person and Group.txt', 'r', encoding='utf8') as f:
    receivers = [receiver.strip() for receiver in  f.readlines()]
print(sys.argv)
print ('So tham so:', len(sys.argv), 'tham so.')
print ('Danh sach tham so:', str(sys.argv))
dt = datetime.today()
seconds = dt.timestamp()
print(seconds)

a = os.path.join(currentdir, "Monitoring_Group")
print(a)

groups_list = {}
with open(os.path.join(currentdir, "Monitoring_Group"), 'r', encoding='utf8') as f:
    groups_list_container = {
        _.split(":")[0].strip(): [
            g.strip()
            for g in _.split(":")[1].strip().split(";")
            if len(g) > 0
        ]
        for _ in f.read().strip().split("#")
        if len(_) > 0
    }
    keys = groups_list_container.keys()
    print(keys)
    for k in keys:
        groups_list[k] = {
            "Whatsapp":[_
                for _ in groups_list_container[k]
                if not ('@firstsolar.com' in _)
            ],
            "Email":[_
                for _ in groups_list_container[k]
                if '@firstsolar.com' in _
            ]
        }
print (groups_list_container[k])
print (groups_list[k])
print (groups_list['Giang'])
print(groups_list_container)
a = 'EdgeSealInspectionDispense:Finishing Monitoring;#'
print(len(a))

def DataCollection(DataName, SelectedSite, StartTime, EndTime):
    
    server = f"Driver=SQL Server;Server={SelectedSite}MESSQLODS;Database=ODS;Trusted_Connection=Yes"
    datasource = sql_sourcecode_list[DataName+'SQL']\
                .replace("__SelectedSite__", SelectedSite)\
                .replace("__StartTime__", StartTime)\
                .replace("__EndTime__", EndTime)
    connection = pyodbc.connect(server)
    connection.cursor()
    #print(datasource)
    DataCollected = pd.read_sql_query(datasource, connection, index_col=None)
    print((DataCollected.index))
    print(type(DataCollected))
    return DataCollected

def DataManiulating(DataCollected, DataName):
    google_drive_folder = google_drive_path
    DataStorage = location_drive_dict[DataName + "Storage"].replace("__Google_Drive_Folder__,", google_drive_folder)
    if os.path.exists(DataStorage):

drive_path = "C:\My Drive"
DataName = "EdgeSealInspection"
drive_folder = drive_path
DataStorage = location_drive_dict[DataName + "Storage"].replace("__Drive_Folder__", drive_path)

if os.path.isdir(DataStorage):
    print(f"{DataName} Storage Has Been Existed : {DataStorage}")
else:
    print(f"{DataName} Has Not Existed => Create A New Folder To Save Data")
    os.mkdir(DataStorage)

# *arg - non keyword argument
def myfun1(*argv):
    for arg in argv:
        print(arg)
myfun1("Hello", "Nguyen", "Văn", "Giang")
def myfun2(arg1, *argv):
    print("Fist Argument : ", arg1)
    for arg in argv:
        print("Next Argument Through * argv : ", arg)
myfun2("Hello", "Nguyen", "Văn", "Giang")
# **kwag - keyword argument
dict = {"a":1, "b":2, "c":3, "d":4}
print(dict)
print(dict.items())
def myfun3(**kwargs):
    for key in kwargs.keys():
        print(key)
myfun3(a=1, b=2)
'''

'''
google_drive_path = "C:\My Drive"
DataName = "EdgeSealInspection"
drive_folder = drive_path
DataStorage = location_drive_dict[DataName + "Storage"].replace("__Drive_Folder__", drive_path)
if os.path.isdir(DataStorage):
    print(f"{DataName} Storage Has Been Existed : {DataStorage}")
else:
    print(f"{DataName} Has Not Existed => Create A New Folder To Save Data")
    os.mkdir(DataStorage)

def Manipulating(DataCollected, DataName, **kwargs):
    # Prevent multi varible has name is same and impact to name of file saved
    datasaving = os.path.join(DataStorage, DataName + f"{kwargs['ext_name'] if 'ext_name' in kwargs.keys() else ''}.csv") 
    root, *_ = os.walk(DataStorage)
    img_list = [_ for _ in root[-1] if _.endswith("png") or _.endswith("jpg") or _.endswith("jpeg")]
    for img in img_list:
        os.remove(os.path.join(DataStorage, img))
    if DataName + f"{kwargs['ext_name'] if 'ext_name' in kwargs.keys() else '' }.csv" in root[-1]:
        os.remove(os.path.join(DataStorage, DataName + f"{kwargs['ext_name'] if 'ext_name' in kwargs.keys() else ''}.csv"))
    img_list = []
    if len(DataCollected.index) == 0:
        pass
    else:
        DataCollected.to_csv(datasaving, index=False)
    # Call a subprocess and generate the monitoring charts

DataCollected = DataCollection('EdgeSealInspection', 'DMT2', '2021-08-09 02:00:00', '2021-08-09 04:00:00')
Manipulating(DataCollected = DataCollected, DataName='EdgeSealInspection', ext_name = '_GlassToBead')

root, *_ = os.walk(DataStorage)
print(root)

def f():
    return 1,2
def g():
    return 1,2,3,
def h():
    return 1,2,3,4

x,y,z,*_ = g()
print(_)
x,y,z,*_ = h()
print(_)

for(root,dir,file) in os.walk('C:\\Users\\fs120806\\Desktop\\Auto Sending'):
    print(root)
    print("____________________")

root, *_ = os.walk('C:\\Users\\fs120806\\Desktop\\Auto Sending')
print(root[-1])
# mỗi lần trả về 3 thông số, tupple
# root[-1] sẽ là lấy file trong tupple [root, dir, file]
'''

'''
# root là đường dẫn của thư mục con và thư mục chính
# dir là thư mục không bao gồm thư mục chính, trả về empty nếu đó là thư mục rỗng
# file là tất cả tập tin trong thư mục, không bao gồm thư mục

for a in os.walk('C:\\Users\\fs120806\\Desktop\\Auto Sending', topdown=True):
    print(a)

# mỗi lần duyệt nó sẽ trả về một tuple gồm root, dir, file
# nó sẽ duyệt nhiều lần đến khi nào hết hết file thì thôi
# os.walk trả về dữ liệu kiểu object generator mà phải lặp for mới lấy được dữ liệu

b = os.walk('C:\\Users\\fs120806\\Desktop\\Auto Sending', topdown=True)
print(b)

# Generator chẳng qua cũng chỉ là Iterator, chúng cũng tạo ra một đối tượng kiểu danh sách.
# Bạn chỉ có thể duyệt qua các phần tử của generator một lần duy nhất.
# Vì generator không lưu dữ liệu trong bộ nhớ mà cứ mỗilần lặp thì chúng sẽ tạo phần tử tiếp theo trong dãy và trả về phần tử đó.
'''

'''
datapath = os.path.join(r"__DataLocationPath__", "")
print(datapath)
'''

'''
dt = pd.read_csv('G:\My Drive\Edge_Seal_GlassToBead_Inspection_Storage\EdgeSealInspection_GlassToBead.csv', index_col=None)
print(type(dt))
# Index Col chọn cái nào thì tên đó sẽ dùng làm cột index

subprocess.call(['C:\\Program Files\\SAS\\JMP\\13\\jmp.exe','C:\\Users\\fs120806\\Desktop\\Auto Sending\\Save_JMP_File.jsl'])
time.sleep(1)
print("Save Image Successfully")

# Terminate_Process('Google Chrome')
Terminate_Process('chrome.exe')
Terminate_Process('jmp.exe')

with open(os.path.join(os.path.dirname(__file__), "DataLocationPath"), "r") as datapath:
    location_drive_dict = {
        _.split("=")[0].strip():_.split("=")[1].strip()
            for _ in datapath.read().strip().split(";")
            if len(_) > 0
    }
print(location_drive_dict)


dataname_list = []
with open(os.path.join(os.path.dirname(__file__),"SQL_Source_Code"), "r", encoding='utf-8') as source:
    dataname_list = [
        _.split(":")[0].strip().strip('SQL')
        for _ in source.read().split(";")
        if len(_) > 0
        ]
for i in range(len(dataname_list)):
    print((dataname_list[0])+"Giang")

dataname_list = []
with open(os.path.join(os.path.dirname(__file__),"SQL_Source_Code"), "r", encoding='utf-8') as source:
    dataname_list = [
        _.split(":")[0].strip().split("_")[0].strip('SQL')
        for _ in source.read().split(";")
        if len(_) > 0]
print(dataname_list)

string_mark_time = []
with open(os.path.join(os.path.dirname(__file__),"SQL_Source_Code"), "r", encoding='utf-8') as source:
    string_mark_time = [
        _.split(":")[0].strip().split("_")[1]
        for _ in source.read().split(";")
        if len(_) > 0]
print(string_mark_time)

if((dataname_list[1].split('_')[1]) == '1'):
    print("Giang")
a = ['Giang','Chi','Thang', 'Phuc']
b = ['Giang','Phuc','Trinh','Nhi','Dung','Ngoc','Yen']
c = ['Chi','Phuc','Trinh','Nhi','Dung','Ngoc','Yen']
d = ['Thang','Phuc','Trinh','Nhi','Dung','Ngoc','Yen']
for i in range (len(a)):
    if a[i] in b:
        print('Giang')
    elif a[i] in c:
        print('Chi')
    elif a[i] in d:
        print('Thang')

# Set Condition For Trigger Time
for _ in range (len(dataname_list)):
    if dataname_list[_] in dataname_list_trigger_30:
        TriggerTime = 0.5
    elif dataname_list[_] in dataname_list_trigger_60:
        TriggerTime = 1

dataname_list_trigger_30 = ['EdgeSealInspection']
dataname_list_trigger_60 = ['CoverGlassRobotAlarm']

for i in dataname_list:
    print(i)

def initalizing_chrome():
    start = datetime.now()
    startsecond = 0
    endsecond = 0
    startsecond = start.hour*60*60 + start.minute*60 + start.second + (start.microsecond/1000000)
    url = "https://web.whatsapp.com/"
    chromedirectory = "user-data-dir=C:\\Users\\fs120806\AppData\Local\Google\Chrome\\User Data"
    options = webdriver.ChromeOptions()
    options.add_argument(chromedirectory)
    driver = webdriver.Chrome(executable_path="C:\Chrome Driver\chromedriver.exe", options=options)
    driver.maximize_window()
    print("Start Sending Message")
    driver.get(url)
    search_xpath = '//*[@id="side"]/div[1]/div/label/div/div[2]'
    search_box = WebDriverWait(driver, 40).until(EC.element_to_be_clickable((By.XPATH, search_xpath)))

initalizing_chrome()

subprocess.call(['C:\\Program Files\\SAS\\JMP\\13\\jmp.exe','C:\\Users\\fs120806\\Desktop\\Auto Sending -  Giang F1 - Developing\\EdgeSealInspection_GlassToBead.jsl'])
time.sleep(1)
print("Save Image Successfully")
time.sleep(2)
send_attachment(receiver_group_name, file_path_edge_seal, message = TimeRangeTrigger)
Terminate_Process('chrome.exe')

if len(CoverGlassAlarmData.index) == 0:
    control = 0
    print('Empty Data Was Fetched')
else:
    control = 1
    print('Data Existed')
    subprocess.call(['C:\\Program Files\\SAS\\JMP\\13\\jmp.exe','C:\\Users\\fs120806\\Desktop\\Auto Sending -  Giang F1 - Developing\\CoverGlassRobot_Alarm.jsl'])
    time.sleep(1)
    print("Save Image Successfully")
    time.sleep(2)
send_attachment(receiver_group_name, file_path_cover_glass, message = TimeRangeTrigger, control=control)

a = 'EdgeSealInspection'
jsl_code_file = f'C:\\Users\\fs120806\\Desktop\\Auto Sending -  Giang F1 - Developing\\{a}_Jmp_File.jsl'
print(f'C:\\Program Files\\SAS\\JMP\\13\\jmp.exe', jsl_code_file)

DataCollected.to_csv(r'DataStorage.csv', index=False)

def func(a,b):
    print(a+b)

url = "https://web.whatsapp.com/"
chromedirectory = "user-data-dir=C:\\Users\\fs120806\AppData\Local\Google\Chrome\\User Data"
options = webdriver.ChromeOptions()
options.add_argument(chromedirectory)
driver = webdriver.Chrome(executable_path="C:\Chrome Driver\chromedriver.exe", options=options)
driver.maximize_window()
print("Start Sending Message")
driver.get(url)
print('Giang')
#func(2,3)

drive_path = "C:\My Drive"

dataname_list = []
with open(os.path.join(os.path.dirname(__file__),"SQL_Source_Code"), "r", encoding='utf-8') as source:
    dataname_list = [
        _.split(":")[0].strip().split("_")[0].strip('SQL')
        for _ in source.read().split(";")
        if len(_) > 0]

location_drive_dict = []
with open(os.path.join(os.path.dirname(__file__), "DataLocationPath"), "r", encoding='utf-8') as datapath:
    location_drive_dict = {
        _.split("=")[0].strip():_.split("=")[1].strip()
            for _ in datapath.read().strip().split(";")
            if len(_) > 0
    }

print(location_drive_dict[dataname_list[1]])
print('Giang')
print(dataname_list)
print(os.path.join((location_drive_dict[dataname_list[1] + "_IMGFile"].replace("__Drive_Folder__", drive_path)), dataname_list[1] + '.png'))
os.path.join((location_drive_dict[dataname_list[1] + "_IMGFile"].replace("__Drive_Folder__", drive_path)), dataname_list[1] + '.png')

def initalizing_chrome():
    url = "https://web.whatsapp.com/"
    chromedirectory = "user-data-dir=C:\\Users\\fs120806\AppData\Local\Google\Chrome\\User Data"
    options = webdriver.ChromeOptions()
    options.add_argument(chromedirectory)
    initalizing_chrome.driver = webdriver.Chrome(executable_path="C:\Chrome Driver\chromedriver.exe", options=options)
    initalizing_chrome.driver.maximize_window()
    print("Start Sending Message")
    initalizing_chrome.driver.get(url)
initalizing_chrome()

with open(os.path.join(os.path.dirname(__file__), "Sql_Source_Code"), "r") as source:
    dataname_hour = {
        _.split("SQL_")[0].strip():_.split("SQL_")[1].strip().split(":")[0]
        for _ in source.read().split(";")
        if len(_) > 0
    }
print(dataname_hour)

def isdriverclosed():
    try:
        driver.get("Google.com")
        return False
    except WebDriverException:
        print("Webdriver Closed By User")
        return True

def initalizing_chrome():
    url = "https://web.whatsapp.com/"
    chromedirectory = "user-data-dir=C:\\Users\\fs120806\AppData\Local\Google\Chrome\\User Data"
    options = webdriver.ChromeOptions()
    options.add_argument(chromedirectory)
    driver = webdriver.Chrome(executable_path="C:\Chrome Driver\chromedriver.exe", options=options)
    driver.maximize_window()
    print("Start Sending Message")
    driver.get(url)
    time.sleep(20)
initalizing_chrome()

with open(os.path.join(os.path.dirname(__file__), "Sql_Source_Code"), "r") as source:
    dataname_hour = {
        _.split("SQL_")[0].strip():_.split("SQL_")[1].strip().split(":")[0]
        for _ in source.read().split(";")
        if len(_) > 0
    }


print(list(dataname_hour.keys())[0])
print(list(dataname_hour.values())[0])

import psutil

def if_process_is_running_by_exename(exename='chrome.exe'):
    for proc in psutil.process_iter(['pid', 'name']):
        # This will check if there exists any process running with executable name
        if proc.info['name'] == exename:
            return True
    return False

import psutil

def if_process_is_running_by_exename(exename='chrome.exe'):
    for proc in psutil.process_iter(['pid', 'name']):
        # This will check if there exists any process running with executable name
        if proc.info['name'] == exename:
            print('Open')
    return print('Close')
if_process_is_running_by_exename(exename='chrome.exe')

def func(a,b):
    func.d = 4
    print(a + b)
func(2,3)
x = func(2,3)
print(func.d)

url = "https://web.whatsapp.com/"
chromedirectory = "user-data-dir=C:\\Users\\fs120806\AppData\Local\Google\Chrome\\User Data"
options = webdriver.ChromeOptions()
options.add_argument(chromedirectory)
driver = webdriver.Chrome(executable_path="C:\Chrome Driver\chromedriver.exe")

c = 1 
def func():
    if c == 0:
        print('Giang')
    else:
        print('Dung')
    print('Giang')
func()

# Initalizin Google Chrome And Access To Whatsapp
initalizing_chrome()

# Process To Trigger Message To Whattsapp
def sending_multidata_to_multigroup(dataname_and_trigger_time, receiver_and_group_name):
    for i in range(len((dataname_and_trigger_time))):
        print(f'Start Getting Data Of {list(dataname_and_trigger_time.keys())[0]}')
        triggertime = float(list(dataname_and_trigger_time.values())[i])
        Initial_DateTime = datetime.today().replace(minute=0, second=0)
        LastTriggerTime = Initial_DateTime - timedelta(hours=triggertime)
        TimeRangeTrigger = f'Data Trigger From {LastTriggerTime.strftime("%d-%m-%Y %H:%M:%S %p")} To {Initial_DateTime.strftime("%d-%m-%Y %H:%M:%S %p")}'
        print(f'Start Time Trigger: {Initial_DateTime.strftime("%d-%m-%Y %H:%M:%S %p")}')
        print(f'Start Running Time: {Start.strftime("%d-%m-%Y %H:%M:%S %p")}')
        print(f'Last Trigger Time: {LastTriggerTime.strftime("%d-%m-%Y %H:%M:%S %p")}')
        databridge = DataCollection(
            list(dataname_and_trigger_time.keys())[i],
            "DMT2",
            (Initial_DateTime - timedelta(hours=triggertime)).strftime('%Y-%m-%d %H:%M:%S'),
            Initial_DateTime.strftime('%Y-%m-%d %H:%M:%S')
        )
        Manipulating(
        databridge, 
        DataName = list(dataname_and_trigger_time.keys())[i], 
        ext_name = '_Csv_File'
        )
        if len(databridge.index) >= 1:
            print(f'Fetched Data Of {list(dataname_and_trigger_time.keys())[i]} To Csv File Successfully') 
            jsl_code_file = f'C:\\Users\\fs120806\Desktop\Auto Sending Developing\\{list(dataname_and_trigger_time.keys())[i]}_Jmp_File.jsl'
            subprocess.call([f'C:\\Program Files\\SAS\\JMP\\16\\jmp.exe', jsl_code_file])
            send_attachment(receiver_and_group_name, image_file=
            os.path.join((location_drive_dict[list(dataname_and_trigger_time.keys())[i] + "_IMGFile"].replace("__Drive_Folder__", drive_path)), list(dataname_and_trigger_time.keys())[i] + '.png'),
            message=TimeRangeTrigger)   
        else:
            print(f'Data Of {dataname_and_trigger_time.keys()[i]} Is Empty')

sending_multidata_to_multigroup(dataname_hour, receiver_group_name)

with open(os.path.join(os.path.dirname(__file__), "Sql_Source_Code"), "r") as source:
    dataname_hour = {
        _.split("SQL_")[0].strip():_.split("SQL_")[1].strip().split(":")[0]
        for _ in source.read().split(";")
        if len(_) > 0
    }
print(dataname_hour[1])

def change():
    for i in range(5):
        if i == 2:
            k = 2
            print(k)
        elif i == 3:
            k = 3
            print(k)
change()
'''