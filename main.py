import os
import glob
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import time
import extract_ncms_info
import excel_manipulation
 
# open HTML page with simulated board
html_driver = webdriver.Chrome('C:/Users/B0640469/Documents/webdriver/chromedriver.exe')
html_driver.get("file:///C:/Users/B0640469/Downloads/NCMS.htm")
#action = ActionChains(html_driver)
#action.send_keys(Keys.F11).perform()
 
while True:
 
    driver = webdriver.Chrome('C:/Users/B0640469/Documents/webdriver/chromedriver.exe')
 
    # extract NCMS information to an HTML file without disposition text
    extract_ncms_info.extract_info(driver, False, 70129)
 
    # extract NCMS information to an HTML file with disposition text
    extract_ncms_info.extract_info(driver, True, 70129)
 
    time.sleep(10)
 
    # find latest file (with dispo) in downloads folder
    files = glob.glob('C:/Users/B0640469/Downloads/*')
    latest_with_dispo = max(files, key=os.path.getctime)
 
    print(latest_with_dispo)
 
    # read exported HTML file into a dataframe
    df_with_dispo = pd.read_html(latest_with_dispo)
 
    # remove HTML file from downloads
    os.remove(latest_with_dispo)
 
    # find latest file (without dispo) in downloads folder
    files = glob.glob('C:/Users/B0640469/Downloads/*')
    latest_without_dispo = max(files, key=os.path.getctime)
 
    # read exported HTML file into a dataframe
    df_without_dispo = pd.read_html(latest_without_dispo)
 
    # remove HTML file from downloads
    os.remove(latest_without_dispo)
 
    print(df_with_dispo[0].columns.values)
    print(df_with_dispo[0]['Discrepancy Status'])
    print(df_with_dispo[0]['NCR Display Status'])
    print(df_with_dispo[0]['NCR Type'])
 
    print('with dispo: ', df_with_dispo)
    print('without dispo: ', df_without_dispo)
 
    # convert DataFrames to an HTML page
    excel_manipulation.df_to_html(df_with_dispo, df_without_dispo)
 
    # refresh HTML
    html_driver.refresh()
 
    driver.close()
 
    # reload results at specified intervals
    time.sleep(5)