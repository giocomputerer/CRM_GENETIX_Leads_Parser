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


username="annac@genetix.ai"
password="G3nEt!x5002"


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





time.sleep(3)

with open("Leads_to_parse.txt") as f:
    leadsraw = f.readlines()

leads_to_parse=[]

for i in range(len(leadsraw)):
    leads_to_parse.append(leadsraw[i].strip())
   



wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1')


for k in range(int(len(leads_to_parse)/10)):
    l=k*10 +6
    driver.get("https://crm.genetix.ai/Lead/Detail/" + leads_to_parse[l])
    time.sleep(2)
    wb.save('Report7.xls')
    prospect_details = driver.find_elements_by_tag_name('td')
    
    patient_info = driver.find_elements_by_tag_name('td')
    conditions_table = driver.find_element_by_id('PersonalConditionsTable')
    conditions=conditions_table.find_elements_by_tag_name('td')
    
    sheet1.write(l, 0, leads_to_parse[l])
    
    status_table = driver.find_elements_by_class_name("table")

    statuses= status_table[0].find_elements_by_tag_name('td')
    for ps in range(len(statuses)):
        if(ps==3):
            sheet1.write(l, 1, statuses[ps].get_attribute("textContent"))
            #print(statuses[ps].get_attribute("textContent"))
        elif(ps==5):
            sheet1.write(l, 2, statuses[ps].get_attribute("textContent"))
            #print(statuses[ps].get_attribute("textContent"))
        elif(ps==11):
            sheet1.write(l, 3, statuses[ps].get_attribute("textContent"))
            #print(statuses[ps].get_attribute("textContent"))
        
    credentials= status_table[1].find_elements_by_tag_name('td')
    for ps in range(len(credentials)):
        if(ps==1):
            sheet1.write(l, 4, credentials[ps].get_attribute("textContent"))
            #print(credentials[ps].get_attribute("textContent"))
        elif(ps==5):
            sheet1.write(l, 5, credentials[ps].get_attribute("textContent"))
            #print(credentials[ps].get_attribute("textContent"))
        elif(ps==9):
            sheet1.write(l, 6, credentials[ps].get_attribute("textContent"))
            #print(credentials[ps].get_attribute("textContent"))
        elif(ps==3):
            sheet1.write(l, 7, credentials[ps].get_attribute("textContent"))
            #print(credentials[ps].get_attribute("textContent"))
    
    phone= status_table[2].find_elements_by_tag_name('td')
    for ps in range(len(phone)):
        if(ps==1):
            sheet1.write(l, 8, phone[ps].get_attribute("textContent"))
            #print(phone[ps].get_attribute("textContent"))
    
    Emailadress= status_table[3].find_elements_by_tag_name('td')
    for ps in range(len(Emailadress)):
        if(ps==1):
            sheet1.write(l, 9, Emailadress[ps].get_attribute("textContent"))
            #print(Emailadress[ps].get_attribute("textContent"))
    
    physical_adress= status_table[4].find_elements_by_tag_name('td')
    for ps in range(len(physical_adress)):
        if(ps==2):
            sheet1.write(l, 10, physical_adress[ps].get_attribute("textContent"))
            #print(physical_adress[ps].get_attribute("textContent"), end=' ')
        elif(ps==3):
            sheet1.write(l, 11, physical_adress[ps].get_attribute("textContent"))
            #print(physical_adress[ps].get_attribute("textContent"), end=' ')
        elif(ps==4):
            sheet1.write(l, 12, physical_adress[ps].get_attribute("textContent"))
            #print(physical_adress[ps].get_attribute("textContent"), end=' ')
        elif(ps==5):
            sheet1.write(l, 13, physical_adress[ps].get_attribute("textContent"))
            #print(physical_adress[ps].get_attribute("textContent"), end=' ')
        elif(ps==6):
            sheet1.write(l, 14, physical_adress[ps].get_attribute("textContent"))
            #print(physical_adress[ps].get_attribute("textContent"))
        
    
    print(l)
    HBP=False
    HC=False
    HA=False
    HD=False
    HI=False
    Stroke=False
    SOB=False
    DT=False
    
    
    for pi in range(len(conditions)):
        if(pi%3==1):
            if(str(conditions[pi].get_attribute("textContent"))=="High Blood Pressure"):
                   HBP=True
            if(str(conditions[pi].get_attribute("textContent"))=="High Cholesterol"):
                   HC=True
            if(str(conditions[pi].get_attribute("textContent"))=="Heart Attack"):
                   HA=True
            if(str(conditions[pi].get_attribute("textContent"))=="Heart Disease"):
                   HD=True
            if(str(conditions[pi].get_attribute("textContent"))=="Heart Infection"):
                   HI=True
            if(str(conditions[pi].get_attribute("textContent"))=="Stroke"):
                   Stroke=True
            if(str(conditions[pi].get_attribute("textContent"))=="Shortness of Breath"):
                   SOB=True
            if(str(conditions[pi].get_attribute("textContent"))=="Diabetes Type 2"):
                   DT=True
    
    if (HBP):
        sheet1.write(l, 15, "YES")
    else:
        sheet1.write(l, 15, "NO")
    if (HC):
        sheet1.write(l, 16, "YES")
    else:
        sheet1.write(l, 16, "NO")
    if (HA):
        sheet1.write(l, 17, "YES")
    else:
        sheet1.write(l, 17, "NO")
    if (HD):
        sheet1.write(l, 18, "YES")
    else:
        sheet1.write(l, 18, "NO")
    if (HI):
        sheet1.write(l, 19, "YES")
    else:
        sheet1.write(l, 19, "NO")
    if (Stroke):
        sheet1.write(l, 20, "YES")
    else:
        sheet1.write(l, 20, "NO")
    if (SOB):
        sheet1.write(l, 21, "YES")
    else:
        sheet1.write(l, 21, "NO")
    if (DT):
        sheet1.write(l, 22, "YES")
    else:
        sheet1.write(l, 22, "NO")
        
    
    
    
    
    
    # print(str(prospect_details[1].text)+" ", end="") #phone
    # print(str(prospect_details[5].text)+" ", end="") #Date of Birth
    # print(str(prospect_details[17].text)+" ", end="") # High Blood Pressure
    # print(str(prospect_details[19].text)+" ", end="") # High Cholesterol
    # print(str(prospect_details[21].text)+" ", end="") # Heart Attack
    # print(str(prospect_details[23].text)+" ", end="") # Stroke
    # print(str(prospect_details[25].text)+" ", end="") # Cancer
    # print(str(prospect_details[27].text)+" ") # JP
    
    
    
    




driver.close()