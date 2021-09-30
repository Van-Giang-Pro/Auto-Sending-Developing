from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from datetime import datetime
from selenium.common.exceptions import TimeoutException
import time 
import pyperclip
import timeit
import subprocess
from Config import *

def send_attachment(receiver_group_name, image_file, message):
    for i in range(len(receiver_group_name)):
        search_xpath = '//*[@id="side"]/div[1]/div/label/div/div[2]'
        search_box = WebDriverWait(initalizing_chrome.driver, 40).until(EC.element_to_be_clickable((By.XPATH, search_xpath)))
        search_box.clear()
        search_box.click()
        search_box.send_keys(receiver_group_name[i])
        time.sleep(3)
        print("Sending Message To {}".format(receiver_group_name[i]))
        for retry in range(3):
            try:
                receiver_title = WebDriverWait(initalizing_chrome.driver, 10).until(EC.element_to_be_clickable((By.XPATH, f'//span[@title="{receiver_group_name[i]}"]')))
                break
            except Exception:
                print("Retry:{} {} Is Not Found In Your Contact List".format(retry,receiver_group_name[i]))
                if retry==2:return # return used to end of the function and not continue the function after
        time.sleep(1)
        receiver_title.click()
        input = initalizing_chrome.driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div/div/div[2]/div[1]/div/div[2]')
        input.click()
        input.send_keys(message)
        time.sleep(3)
        input.send_keys(Keys.ENTER)
        time.sleep(3)
        attachment_box = WebDriverWait(initalizing_chrome.driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//div[@title = "Đính kèm"]')))
        time.sleep(3)
        attachment_box.click()
        attachment = initalizing_chrome.driver.find_element_by_xpath('//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
        attachment.send_keys(image_file)
        time.sleep(3)
        send = WebDriverWait(initalizing_chrome.driver, 3).until(EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="send"]')))
        send.click()
        time.sleep(5) # waiting image loaded after sending     
