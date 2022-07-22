#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      R252202
#
# Created:     20/06/2022
# Copyright:   (c) R252202 2022
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import os, shutil
import time
from PIL import ImageGrab

def start(desired_path):

    countries = ["Poland", "Netherlands", "Belgium", "Germany", "Spain", "France", "United Kingdom", "Italy",
    "Portugal", "Sweden", "Turkey"]
    pdfs_files = ["accrual", "audit"]
    jpgs = ["AccrualParams.png", "AuditParams.png"]

    def getScreenshot(path):
        img = ImageGrab.grab()
        img.save(path)
        time.sleep(2)

    def pdf2jpeg(pdf_input_path, jpeg_output_path):

        file = os.startfile(pdf_input_path)
        time.sleep(2)
        getScreenshot(jpeg_output_path)

    import pyexcel as p
    def xlsToXlsx(xls_input_path):
        p.save_book_as(file_name=xls_input_path,
                dest_file_name=xls_input_path + 'x')

    
    file_index = 1
    for country in countries:

        i = 0
        for pdf in pdfs_files:

            try:
                pdf2jpeg(r"I:/ICS/IC/Service Now Reconciliation/temp/Generation/" + str(file_index) + "." + str(country) + r"/" + str(pdf + ".pdf"), r"I:/ICS/IC/Service Now Reconciliation/temp/Generation/"+ str(file_index) + "."  + str(country) + "/" + str(jpgs[i])  )

                xlsToXlsx(r"I:/ICS/IC/Service Now Reconciliation/temp/Generation/"+ str(file_index) + "."  +  str(country) + r"/accrual.xls")
                xlsToXlsx(r"I:/ICS/IC/Service Now Reconciliation/temp/Generation/" + str(file_index) + "." +  str(country) + r"/audit.xls")
                i+= 1
            except:
                print(str(pdf))
        file_index += 1
    
    src = r"I:\ICS\IC\Service Now Reconciliation\temp\Dummy Files\\"
    dst = r"I:\ICS\IC\Service Now Reconciliation\temp\Generation\\"
    print(desired_path)
    shutil.copytree(src, dst, dirs_exist_ok=True)
    shutil.copytree(dst, desired_path, dirs_exist_ok=True)



#.............

if __name__ == '__main__':

    file_index = 1
    for country in countries:

        i = 0
        for pdf in pdfs_files:
            try:
                pdf2jpeg(r"I:/ICS/IC/Service Now Reconciliation/temp/Generation/" + str(file_index) + "." + str(country) + r"/" + str(pdf + ".pdf"), r"I:/ICS/IC/Service Now Reconciliation/temp/Generation/"+ str(file_index) + "."  + str(country) + "/" + str(jpgs[i])  )

                xlsToXlsx(r"I:/ICS/IC/Service Now Reconciliation/temp/Generation/"+ str(file_index) + "."  +  str(country) + r"/accrual.xls")
                xlsToXlsx(r"I:/ICS/IC/Service Now Reconciliation/temp/Generation/" + str(file_index) + "." +  str(country) + r"/audit.xls")
            except:
                print(str(pdf))
            i+= 1
        file_index += 1


