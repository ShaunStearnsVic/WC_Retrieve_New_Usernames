#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import os
import time
import keyring
import openpyxl
from openpyxl import load_workbook
import pandas as pd
import csv

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.common.exceptions import (
    NoSuchElementException,
    ElementNotVisibleException,
)
import chromedriver_binary
from webdriver_manager.chrome import ChromeDriverManager


# ## Required Libraries
# #### os (Installed with python)
# #### time (Installed with python)
# #### keyring (conda install -c anaconda keyring)
# #### openpyxl (conda install -c anaconda openpyxl)
# #### pandas (conda install pandas)
# #### selenium (conda install -c anaconda selenium)
# #### chromedriver_binary (conda install -c conda-forge python-chromedriver-binary=87)
#    ###### NOTE: Replace "=87" with whatever version of Chrome you have running. Don't include numbers after first decimal.
# #### webdriver_manager (pip install webdriver_manager)

# # See Names of Registered Users per Facility (First Row Only)

# In[1]:


# function to take care of downloading file
def enable_download_headless(browser, download_dir):
    browser.command_executor._commands["send_command"] = (
        "POST",
        "/session/$sessionId/chromium/send_command",
    )
    params = {
        "cmd": "Page.setDownloadBehavior",
        "params": {"behavior": "allow", "downloadPath": download_dir},
    }
    browser.execute("send_command", params)


# instantiate a chrome options object so you can set the size and headless preference
# some of these chrome options might be uncessary but I just used a boilerplate
# change the <path_to_download_default_directory> to whatever your default download folder is located
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--window-size=1920x1080")
chrome_options.add_argument("--disable-notifications")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--verbose")
chrome_options.add_experimental_option(
    "prefs",
    {
        "download.default_directory": "<path_to_download_default_directory>",
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing_for_trusted_sources_enabled": False,
        "safebrowsing.enabled": False,
    },
)
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--disable-software-rasterizer")

# initialize driver object and change the <path_to_chrome_driver> depending on your directory where your chromedriver should be
driver = webdriver.Chrome()

# Portal Page
driver.get("https://www.dcphrapps.dphe.state.co.us/Account/Login")

###################################################################
# Insert Your Log In Credentials
###################################################################
#Log in
username = driver.find_element_by_name("Email")
username.clear()
username.send_keys("USERNAME/EMAIL")

password = driver.find_element_by_name("Password")
password.clear()
password.send_keys("YOUR PASSWORD")
##################################################################

driver.find_element_by_css_selector('[value="Log in"]').click()
time.sleep(1)

# Click “Sites”
try:
    driver.find_element_by_xpath('//*[@id="SitesButton"]/span[1]').click()
    time.sleep(5)
except (ElementNotVisibleException, NoSuchElementException):
    driver.find_element_by_id("DashButton").click()
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="ApplicationButton_45"]/span').click()
    time.sleep(5)
    pass

# Select "Show All"
show_all = Select(driver.find_element_by_name("SiteList_length"))
show_all.select_by_visible_text("All")
time.sleep(1)

##########################################################################
#Change File path to Location of Facility Names.xlsx
##########################################################################
sites = pd.read_excel(r"C:\.....\Facility Names.xlsx")
list = sites["Sites"]
##########################################################################

# Make windowfull screen
driver.maximize_window()

## Above cell must first be run
for i in list:
    driver.find_element_by_css_selector("input[type='search']").send_keys(i)
    time.sleep(1)
    try:
        vals = driver.find_elements_by_xpath('//*[@id="SiteList"]/tbody//tr[1]/td[4]')
        for val in vals:
            print(i,",       ",val.text)
    except:
        print(i,",       ", "No Users")
        pass
    driver.find_element_by_css_selector("input[type='search']").clear()
    time.sleep(1)

