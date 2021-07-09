from selenium import webdriver
from selenium.webdriver.common.by import By
import io
import time
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import xlwt
from xlwt import Workbook
from inspect import getmembers, isfunction





username="annac@genetix.ai"
password="G3nEt!x5002"


pages_to_scrub= input("how many pages should i go through? :")

number_of_good_prospects=0
number_of_desired_prospects=0
leads_to_parse=[]

url = "https://crm.genetix.ai/"
driver = webdriver.Chrome()
driver.get(url)

time.sleep(3)



usernamebar = driver.find_element_by_id('Email')
usernamebar.send_keys(username)

passwordbar = driver.find_element_by_id('Password')
passwordbar.send_keys(password)

nextButton = driver.find_element_by_class_name('btn')
nextButton.click()

driver.get("https://crm.genetix.ai/Lead/MyPipeline?status=24&testCode=100");

time.sleep(3)

amountthing = driver.find_element_by_name('LeadsListTable_length')
needclick=amountthing.find_elements_by_tag_name("option")
amountthing.click()
needclick[5].click()

f=open("Leads_to_parse.txt","a+")

time.sleep(4)

for p in range(int(pages_to_scrub)):  


    datastuff = driver.find_elements_by_tag_name('td')


    for i in range(len(datastuff)):
        if(i%8==0):
            print(datastuff[i].text)
            f.write(str(datastuff[i].text))
            f.write("\n")
            leads_to_parse.append(str(datastuff[i].text))
        
    if((p+1)!= int(pages_to_scrub)):
        next_page=driver.find_element_by_id("LeadsListTable_next")
        next_page.click()
    
    
    print("page", end=" ")
    print("out of"+pages_to_scrub, end=" ")
    
    print(p+1)
    
    time.sleep(4)
    
f.close()




    
    
driver.close()


