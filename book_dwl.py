#!/usr/bin/env python3

import os
import sys
import requests
import json
import csv
import html
import re
import time

from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait

chrome_wd_path = r"C:\Users\ivaign\OneDrive - United Conveyor Corp\Documents\Python_Projects\chrome-win64\chrome.exe"


def download_book(book_url):
    print('Trying to download a book ' + book_url)
    current_directory = os.getcwd()
    download_directory = os.path.join(current_directory, 'downloads')

    options = webdriver.ChromeOptions()
    options.binary_location = chrome_wd_path


    driver = webdriver.Chrome(options=options)

    driver.get(book_url)
    # ivanignatenko@uccenvironmental.com
    driver.fullscreen_window()

    time.sleep(5)
    element = driver.find_element(by=By.CLASS_NAME, value="lnXdpd")
    # element = driver.find_element(By.XPATH, r"/html/body/div[3]/div[1]/div/div[3]")
    # element = driver.find_element(by=By.CLASS_NAME, value="encription_password_input_cell")
    # element.click()
    # element.send_keys("ivanignatenko@uccenvironmental.com")

)




    # element = driver.find_element(by=By.XPATH, value=r"/html/body/div[2]/div[1]/div[1]/div[4]/div/div[1]/div/img")
    # print(element)
    # element.screenshot("element_screenshot.png")
    #
    # time.sleep(1)
    #
    # driver.quit()

# download_book(r"https://fliphtml5.com/ufdzy/rvgw/")
download_book(r"https://www.google.com/")

#