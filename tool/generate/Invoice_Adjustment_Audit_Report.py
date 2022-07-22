#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      R252202
#
# Created:     18/05/2022
# Copyright:   (c) R252202 2022
# Licence:     <your licence>
#-------------------------------------------------------------------------------
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import keyboard
import time
import shutil
import os
from PIL import ImageGrab

# pip install webdriver-manager
# pip install selenium

# https://ppg.service-now.com/financepulse?id=credit_audit_report
# https://ppg.service-now.com/financepulse?id=credit_accrual_report

# TODO:
# Download PDF
# Manage downloads
# Screenshots
# Naming the files

# All these files in one folder with proper name




def start(period):
    countries = ["Poland", "Netherlands", "Belgium", "Germany", "Spain", "France", "United Kingdom", "Italy",
    "Portugal", "Sweden", "Turkey"]

    #download_path = 'I:\\ICS\\IC\\Service Now Reconciliation\\temp\\Downloads'
    download_path = getDownloadPath()

    if os.path.exists(download_path + r"\u_customer_credit_memo.xls"):
        os.remove(download_path + r"\u_customer_credit_memo.xls")

    if os.path.exists(download_path + r"\u_customer_credit_memo.pdf"):
        os.remove(download_path + r"\u_customer_credit_memo.pdf")

    chrome_options1 = webdriver.ChromeOptions()
    prefs = {'download.default_directory' : download_path}
    chrome_options1.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options = chrome_options1)

    driver.get('https://ppg.service-now.com/financepulse?id=credit_audit_report')

    element =driver.find_element_by_xpath("//span[@class='largeTextNoWrap indentNonCollapsible']")
    element.click()


    def setup(period):

        #REQUEST TYPE
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//span[@class='select2-arrow']"))).click()
        time.sleep(2)
        keyboard.write("")
        time.sleep(1)
        keyboard.press_and_release('esc')

        #SBU
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//input[@class='select2-input select2-default']"))).click()
        time.sleep(2)
        keyboard.write("Industrial")
        time.sleep(1)
        keyboard.press_and_release('enter')

        #REGION
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//div[@class='select2-container select2-container-multi ng-pristine ng-untouched ng-valid ng-isolate-scope ng-empty select2-reference ng-form-element']"))).click()
        time.sleep(2)
        keyboard.write("Europe")
        time.sleep(1)
        keyboard.press_and_release('enter')

        #PERIOD
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//div[@id='s2id_autogen9']"))).click()
        time.sleep(2)
        keyboard.write(period)
        time.sleep(1)
        keyboard.press_and_release('enter')

    def getScreenshot(path):
        img = ImageGrab.grab()
        img.save(path + "params_audit.png")
        time.sleep(2)



    def change_country(countries, id):
        i=1
        for country in countries:
            #COUNTRY
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//div[@id='s2id_autogen7']"))).click()
            time.sleep(2)
            keyboard.write(country)
            time.sleep(1)
            keyboard.press_and_release('down')
            keyboard.press_and_release('enter')

            time.sleep(3)
            download_form()
            time.sleep(7)

            moveAndRenameFile(download_path, "/u_customer_credit_memo.xls", r"I:/ICS/IC/Service Now Reconciliation/temp/Generation/"+ str(i) + ".", str(country), ".xls")
            moveAndRenameFile(download_path, "/u_customer_credit_memo.pdf", r"I:/ICS/IC/Service Now Reconciliation/temp/Generation/"+ str(i) + ".", str(country), ".pdf")
            #moveAndRenameFile(r"//EUC.PPG.COM/dfs/WRC/users/" + str(id) + "/", "u_customer_credit_memo.xls", "I:\\ICS\\IC\\Service Now Reconciliation\\temp\\Generation\\" + str(i) + ".",  country, ".xls")
            #moveAndRenameFile(r"//EUC.PPG.COM/dfs/WRC/users/"  + str(id) + "/", "u_customer_credit_memo.pdf", "I:\\ICS\\IC\\Service Now Reconciliation\\temp\\Generation\\" + str(i) + ".",  country, ".pdf")
            #getScreenshot(r"I:/ICS/IC/Service Now Reconciliation/temp/Generation/" + str(country) + "/")
            #moveAndRenameFile("\\EUC.PPG.COM\dfs\WRC\users\R252202", "u_customer_credit_memo.pdf", "I:\\ICS\\IC\\Service Now Reconciliation\\temp\\" + str(country) + "\\", country)
            getScreenshot(r"I:/ICS/IC/Service Now Reconciliation/temp/Generation/" + str(i) + "." + str(country) + "/")


            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//div[@id='s2id_autogen7']"))).click()
            keyboard.press_and_release('backspace')
            keyboard.press_and_release('backspace')
            time.sleep(3)
            i+=1

    def download_form():
        #APPLY
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//button[@name='u_apply_filter']"))).click()
        time.sleep(3)
        #MENU
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//span[@class='dropdown m-r-xs']"))).click()

        #DOWNLOAD EXCEL
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//a[text()='Export as Excel']"))).click()

        #MENU
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//span[@class='dropdown m-r-xs']"))).click()

        #DOWNLOAD PDF
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//a[text()='Export as PDF']"))).click()


    def moveAndRenameFile(path_from, file, path_to, country, type):
        try:
            what_to_replace = str(path_from) + str(file)
            print(what_to_replace)
            where_to_replace = str(path_to) + str(country) + "/" + "audit" + str(type)
            print(where_to_replace)
            #shutil.move(what_to_replace, where_to_replace)
            try:
                #path = os.path.join(path_to, str(country))
                path = path_to + country
                print(path)
                os.mkdir(path)
            except:
                print("Folder already exists")
            shutil.move(what_to_replace, where_to_replace)
        except FileNotFoundError as error:
            print(error)

    setup(period)
    change_country(countries, id)




def getDownloadPath():
    #SETUP THE BROWSER
    driver = webdriver.Chrome(ChromeDriverManager().install())

    driver.get('chrome://settings/downloads')
    driver.maximize_window()
    driver.implicitly_wait(2)

    def expand_shadow_element(element):
      shadow_root = driver.execute_script('return arguments[0].shadowRoot', element)
      return shadow_root

    root1 = driver.find_element_by_tag_name('settings-ui')
    shadow_root1 = expand_shadow_element(root1)

    root2 = shadow_root1.find_element(By.CSS_SELECTOR, 'settings-main')
    shadow_root2 = expand_shadow_element(root2)

    root3 = shadow_root2.find_element(By.CSS_SELECTOR,'settings-basic-page')
    shadow_root3 = expand_shadow_element(root3)


    root4 = shadow_root3.find_element(By.CSS_SELECTOR,'settings-downloads-page')
    shadow_root4 = expand_shadow_element(root4)

    element = shadow_root4.find_element(By.CSS_SELECTOR, "div#defaultDownloadPath.secondary")
    element.click()

    print(element.text)
    return element.text

if __name__ == "__main__":
    start("JUN-22", "R252202")


