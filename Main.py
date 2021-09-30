from Database import Manipulating
from Database import DataCollection
from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from datetime import date, datetime, timedelta
from Config import *
import subprocess
import time
from GroupSending import *

Start = datetime.today()

# Caculating Time Elapsed, Starting Time Point
start = datetime.now()
startsecond = 0
endsecond = 0
startsecond = start.hour*60*60 + start.minute*60 + start.second + (start.microsecond/1000000)

# Close task Chrome and JMP before running
Terminate_Process('chrome.exe')
Terminate_Process('jmp.exe')
print("Kill Process Successfully")

# Use Variable To Initalize Chrome At First Time Running
k = 0

# Function To Send Whatsapp Message Trigger
def sending_multidata_to_multigroup(dataname_and_trigger_time, receiver_and_group_name):
    global k
    global driver # driver is created once times at first time if data existed, keep driver is varible is global to use for some functions after
    for i in range(len((dataname_and_trigger_time))):
        print(f'Start Getting Data Of {list(dataname_and_trigger_time.keys())[i]}')
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
            if k == 0:
                url = "https://web.whatsapp.com/"
                chromedirectory = "user-data-dir=C:\\Users\\fs120806\AppData\Local\Google\Chrome\\User Data"
                options = webdriver.ChromeOptions()
                options.add_argument(chromedirectory)
                options.add_experimental_option("detach", True)
                # if false, chrome is quit when driver is killed
                # if true, chrome is only quited when sension is quit or closed
                driver = webdriver.Chrome(executable_path="C:\Chrome Driver\chromedriver.exe", options=options)
                print("Start Sending Message")
                driver.get(url)
                for j in range(len(receiver_and_group_name)):
                    search_xpath = '//*[@id="side"]/div[1]/div/label/div/div[2]'
                    search_box = WebDriverWait(driver, 40).until(EC.element_to_be_clickable((By.XPATH, search_xpath)))
                    search_box.clear()
                    search_box.click()
                    search_box.send_keys(receiver_group_name[j])
                    time.sleep(3)
                    print("Sending Message To {}".format(receiver_and_group_name[j]))
                    for retry in range(3):
                        try:
                            receiver_title = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, f'//span[@title="{receiver_and_group_name[j]}"]')))
                            break
                        except Exception:
                            print("Retry:{} {} Is Not Found In Your Contact List".format(retry,receiver_and_group_name[j]))
                            if retry==2:return # return used to end of the function and not continue the function after
                    time.sleep(1)
                    receiver_title.click()
                    input = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div/div/div[2]/div[1]/div/div[2]')
                    input.click()
                    input.send_keys(TimeRangeTrigger)
                    time.sleep(3)
                    input.send_keys(Keys.ENTER)
                    time.sleep(3)
                    attachment_box = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//div[@title = "Đính kèm"]')))
                    time.sleep(3)
                    attachment_box.click()
                    attachment = driver.find_element_by_xpath('//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
                    attachment.send_keys(os.path.join((location_drive_dict[list(dataname_and_trigger_time.keys())[i] + "_IMGFile"].replace("__Drive_Folder__", drive_path)), list(dataname_and_trigger_time.keys())[i] + '.png'))
                    time.sleep(3)
                    send = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="send"]')))
                    send.click()
                    time.sleep(5) # waiting image loaded after sending 
                    k = k + 1   
            else:
                for j in range(len(receiver_and_group_name)):
                    search_xpath = '//*[@id="side"]/div[1]/div/label/div/div[2]'
                    search_box = WebDriverWait(driver, 40).until(EC.element_to_be_clickable((By.XPATH, search_xpath)))
                    search_box.clear()
                    search_box.click()
                    search_box.send_keys(receiver_and_group_name[j])
                    time.sleep(3)
                    print("Sending Message To {}".format(receiver_and_group_name[j]))
                    for retry in range(3):
                        try:
                            receiver_title = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, f'//span[@title="{receiver_and_group_name[j]}"]')))
                            break
                        except Exception:
                            print("Retry:{} {} Is Not Found In Your Contact List".format(retry,receiver_and_group_name[i]))
                            if retry==2:return # return used to end of the function and not continue the function after
                    time.sleep(1)
                    receiver_title.click()
                    input = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div/div/div[2]/div[1]/div/div[2]')
                    input.click()
                    input.send_keys(TimeRangeTrigger)
                    time.sleep(3)
                    input.send_keys(Keys.ENTER)
                    time.sleep(3)
                    attachment_box = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//div[@title = "Đính kèm"]')))
                    time.sleep(3)
                    attachment_box.click()
                    attachment = driver.find_element_by_xpath('//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
                    attachment.send_keys(os.path.join((location_drive_dict[list(dataname_and_trigger_time.keys())[i] + "_IMGFile"].replace("__Drive_Folder__", drive_path)), list(dataname_and_trigger_time.keys())[i] + '.png'))
                    time.sleep(3)
                    send = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="send"]')))
                    send.click()
                    time.sleep(5) # waiting image loaded after sending
                    k = k + 1     
        else:
            print(f'Data Of {list(dataname_and_trigger_time.keys())[i]} Is Empty') 
    
sending_multidata_to_multigroup(dataname_hour, receiver_group_name) # sending all data to all receiver
# sending_multidata_to_multigroup({'EdgeSealInspection':'300'}, receiver_group_name) # sending single data to all receiver
# sending_multidata_to_multigroup(dataname_hour, ['Data Trigger Fin 10']) # sending all data to single people
# sending_multidata_to_multigroup({'CoverGlassRobotAlarm': '24'}, ['Data Trigger Fin 9']) # sending single data to single people