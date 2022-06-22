import os
import glob
import pandas as pd
from selenium import webdriver
import time
import extract_ncms_info
import excel_manipulation
 
driver = webdriver.Chrome('C:/Users/B0640469/Documents/webdriver/chromedriver.exe')
 
for i in range(1):
    # extract NCMS information to an HTML file without disposition text
    extract_ncms_info.extract_info(driver, False)
 
    # extract NCMS information to aTSn HTML file with disposition text
    extract_ncms_info.extract_info(driver, True)
 
    time.sleep(10)
 
    # find latest file (with dispo) in downloads folder
    files = glob.glob('C:/Users/B0640469/Downloads/*')
    latest_with_dispo = max(files, key=os.path.getctime)
 
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
 
    print('with dispo: ', df_with_dispo)
    print('without dispo: ', df_without_dispo)
 
    excel_manipulation.modify_excel(df_with_dispo, df_without_dispo)