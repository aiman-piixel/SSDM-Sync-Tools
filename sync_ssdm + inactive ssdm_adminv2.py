#!/usr/bin/env python
# coding: utf-8

# In[17]:


from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import os
import time
import glob2
import pandas as pd
from IPython.display import clear_output
import sys

#adjust path accordingly to OS
loginCSV = r"C:\Users\Aiman\Desktop\ssdm damai jaya 16jul-16aug.csv"
excel_folder = r"C:\Users\Aiman\Desktop\DAMAI JAYA"
#loginCSV = "/Users/piixel/Downloads/ssdm damai jaya 16jul-16aug.csv"
#excel_folder = "/Users/piixel/Downloads/DAMAI JAYA"

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--disable-gpu")

#ssdm category
category = {
    "ssdm1": ["MEMBANTU GURU JIKA ADA MAJLIS",
              "MENOLONG DAN MENGAWAL KELAS SEMASA GURU TIADA",
              "MEMBANTU GURU MENGAMBIL KEHADIRAN PELAJAR",
              "MEMBANTU GURU JIKA MEMERLUKAN BANTUAN",
              "MEMBANTU MENGANGKAT BUKU/BARANG GURU"
    ],
    "ssdm2": ["DATANG AWAL KE SEKOLAH", "BERDISIPLIN", "BERCAKAP DENGAN SOPAN",
              "BERPAKAIAN KEMAS", "MENYAPA/MEMBERI SALAM KEPADA GURU"
    ],
    "ssdm3": ["MENYUSUN KEMAS MEJA", "MEMBUANG DEBU/HABUK DI CERMIN/MEJA/DINDING",
              "MEMBERSIHKAN PAPAN HITAM", "MENGEMOP BILIK", "MENYAPU BILIK-BILIK KHAS"
    ],
    "ssdm4": ["MENYUSUN MEJA/PERABOT", "MENGHIAS BILIK",
              "MEMBERSIHKAN MEJA/PERABOT YANG BERDEBU", "MOP LANTAI",
              "MENYAPU SAMPAH-SARAP DI SEKELILING KELAS"
    ],
    "ssdm5": ["MUDAH BERINTERAKSI ATAU BERGAUL",
              "MENGHORMATI PENDAPAT ORANG LAIN",
              "BIJAK BERKOMUNIKASI UNTUK MENDAPATKAN MAKLUMAT",
              "MENJADI FASILATATOR"
    ],
    "ssdm6": ["MELAPOR KES MASALAH SOSIAL YANG SERIUS",
              "MELAPOR KES MEMBAWA/MENGGUNAKAN BARANG TERLARANG",
              "MELAPOR KES TIDAK BERDISIPLIN",
              "MELAPOR KES LEWAT KE SEKOLAH",
              "MELAPOR KES PONTENG SEKOLAH/KELAS"
    ],
    "ssdm7": ["MENYUSUN DAN MENCANTIK KAN KAWASAN TAMAN/LANDSKAP",
              "MEMOTONG RUMPUT YANG PANJANG",
              "MEMBANTU MENYIRAM POKO BUNGA",
              "MENGUTIP SAMPAH-SARAP YANG ADA",
              "MEMBANTU MENYAPU DAUN-DAUN KERING",
              "MEMBANTU MENYIRAM POKOK BUNGA"
    ],
    "ssdm8": ["MELAKUKAN DENGAN PENUH SEMANGAT",
              "SUKA MENOLONG ORANG LAIN",
              "MEMBANTU MENGEMAS SEMULA BARANG YANG DIGUNAKAN",
              "BEKERJASAMA MENGECAT DINDING",
              "MENUNJUKKAN KREATIVITI"
    ]
}

#to match item to ssdm category
def category_ssdm(amalanBaik):
    for ssdm, item in category.items():
        if amalanBaik in item:
            return ssdm
    return None

#return amalan baik value in dropdown
def dropdown_ssdm(category):
    select_amalanBaik = driver.find_element(By.XPATH, "//select[@name='AB']")
    select = Select(select_amalanBaik)
    if category == "ssdm1":
        select.select_by_value('1')
    elif category == "ssdm2":
        select.select_by_value('8')
    elif category == "ssdm3":
        select.select_by_value('4')
    elif category == "ssdm4":
        select.select_by_value('5')
    elif category == "ssdm5":
        select.select_by_value('7')
    elif category == "ssdm6":
        select.select_by_value('2')
    elif category == "ssdm7":
        select.select_by_value('3')
    elif category == "ssdm8":
        select.select_by_value('6')
    else:
        time.sleep(0.5)
        print("Item not valid")
        return "Invalid"

#return day value in dropdown
def return_day(DaySeries):
    select_date = driver.find_element(By.XPATH, "//select[@name='_b_hari_mula_tk']")
    select = Select(select_date)
    if DaySeries == "01":
        select.select_by_value('1')
    elif DaySeries == "02":
        select.select_by_value('2')
    elif DaySeries == "03":
        select.select_by_value('3')
    elif DaySeries == "04":
        select.select_by_value('4')
    elif DaySeries == "05":
        select.select_by_value('5')
    elif DaySeries == "06":
        select.select_by_value('6')
    elif DaySeries == "07":
        select.select_by_value('7')
    elif DaySeries == "08":
        select.select_by_value('8')
    elif DaySeries == "09":
        select.select_by_value('9')
    else:    
        select.select_by_value(DaySeries)
        
#return month value in dropdown
def return_month(MonthSeries):
    select_month = driver.find_element(By.XPATH, "//select[@name='_b_bulan_mula_tk']")
    select = Select(select_month)
    if MonthSeries == "01":
        select.select_by_value('JAN')
    elif MonthSeries == "02":
        select.select_by_value('FEB')
    elif MonthSeries == "03":
        select.select_by_value('MAR')
    elif MonthSeries == "04":
        select.select_by_value('APR')
    elif MonthSeries == "05":
        select.select_by_value('MAY')
    elif MonthSeries == "06":
        select.select_by_value('JUN')
    elif MonthSeries == "07":
        select.select_by_value('JUL')
    elif MonthSeries == "08":
        select.select_by_value('AUG')
    elif MonthSeries == "09":
        select.select_by_value('SEP')
    elif MonthSeries == "10":
        select.select_by_value('OCT')
    elif MonthSeries == "11":
        select.select_by_value('NOV')
    else:
        select.select_by_value('DEC')

#return year value in dropdown
def return_year(YearSeries):
    select_year = driver.find_element(By.XPATH, "//select[@name='_b_tahun_mula_tk']")
    select = Select(select_year)
    select.select_by_value(YearSeries)

#return hour and AM/PM value in dropdown
def return_hour(HourSeries):
    select_hour = driver.find_element(By.XPATH, "//select[@name='jam']")
    select = Select(select_hour)
    if HourSeries == "13" or HourSeries == "01":
        select.select_by_value('01')
    elif HourSeries == "14" or HourSeries == "02":
        select.select_by_value('02')
    elif HourSeries == "15" or HourSeries == "03":
        select.select_by_value('03')
    elif HourSeries == "16" or HourSeries == "04":
        select.select_by_value('04')
    elif HourSeries == "17" or HourSeries == "05":
        select.select_by_value('05')
    elif HourSeries == "18" or HourSeries == "06":
        select.select_by_value('06')
    elif HourSeries == "19" or HourSeries == "07":
        select.select_by_value('07')
    elif HourSeries == "20" or HourSeries == "08":
        select.select_by_value('08')
    elif HourSeries == "21" or HourSeries == "09":
        select.select_by_value('09')
    elif HourSeries == "22" or HourSeries == "10":
        select.select_by_value('10')
    elif HourSeries == "23" or HourSeries == "11":
        select.select_by_value('11')
    elif HourSeries == "24" or HourSeries == "00":
        select.select_by_value('12')    
    else:
        select.select_by_value(HourSeries)
    
    select_AMPM = driver.find_element(By.XPATH, "//select[@name='am_pm']")
    select = Select(select_AMPM)
    if HourSeries == "12":
        select.select_by_value('PM')
    elif HourSeries == "13":
        select.select_by_value('PM')
    elif HourSeries == "14":
        select.select_by_value('PM')
    elif HourSeries == "15":
        select.select_by_value('PM')
    elif HourSeries == "16":
        select.select_by_value('PM')
    elif HourSeries == "17":
        select.select_by_value('PM')
    elif HourSeries == "18":
        select.select_by_value('PM')
    elif HourSeries == "19":
        select.select_by_value('PM')
    elif HourSeries == "20":
        select.select_by_value('PM')
    elif HourSeries == "21":
        select.select_by_value('PM')
    elif HourSeries == "22":
        select.select_by_value('PM')
    elif HourSeries == "23":
        select.select_by_value('PM')
    else:
        select.select_by_value('AM')
        
#return minute value in dropdown
def return_min(MinSeries):
    select_min = driver.find_element(By.XPATH, "//select[@name='minit']")
    select = Select(select_min)
    select.select_by_value(MinSeries)
    
def close_driver():
    try:
        driver.quit()
    except Exception as e:
        print("An error occurred while closing the WebDriver:", e)
    
#read content in ssdm login csv
df_login = pd.read_csv(loginCSV)
schoolName = df_login['NAME']
ssdmID = df_login['ID']
ssdmPass = df_login['PASS']
schoolEmail = df_login['EMAIL']
startPage = df_login['START']
endPage = df_login['END']
statusPrep = df_login['PREPPED'].to_numpy()
statusDone = df_login['DONE'].to_numpy()

#list out all excel files in excel folder
fileNames = glob2.glob(os.path.join(excel_folder + "/*.csv"))
fileNames.sort()

for i in range(len(schoolName)):

    try:
        ############################set ssdm submission inactive in adminv2################################

        #check prep status
        if statusPrep[i]==1:
            print(schoolName[i] +" has already been prepped")
            pass

        else:
            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 10)
            driver.set_script_timeout(10)
            driver.get("https://adminv2.studentqr.com/login")

            login_box, pass_box = wait.until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR , "input[name='email'], input[name='password']")))
            login_box.send_keys(schoolEmail[i])
            pass_box.send_keys('password')
            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']"))).click()
            #let the page load after signing in
            time.sleep(2)

            #set page range for each school
            start_page = startPage[i]
            end_page = endPage[i]

            #loop through from start page to end page and set submission to inactive
            for page_number in range(start_page, end_page + 1):
                page_url = f"https://adminv2.studentqr.com/ssdm?page={page_number}"
                driver.get(page_url)
                time.sleep(1.5)
                #wait for page to load

                #check all checkboxes
                checkboxes = driver.find_elements(By.XPATH, "//input[@type='checkbox']")
                total_checkboxes = len(checkboxes)
                for check_number in range(total_checkboxes):
                    driver.execute_script("arguments[0].click();", checkboxes[check_number])

                #set item to inactive
                select_element = driver.find_element(By.XPATH, "//select[@name='statusCheckbox']")
                select = Select(select_element)
                select.select_by_value('2')

                #update status
                wait.until(EC.element_to_be_clickable((By.XPATH ,"//button[@id='updateStatusButton']"))).click()
                time.sleep(1)
                alert = wait.until(EC.alert_is_present())
                if alert:
                    alert.accept()
                driver.switch_to.default_content()
                time.sleep(4)

                #check if start and end page is the same, end loop if same
                if start_page==end_page:
                    break

            #update school prep status
            statusPrep[i]=1
            df_login['PREPPED'] = statusPrep
            df_login.to_csv(loginCSV, index=False)

            driver.close()
            time.sleep(2)

        ###############################submit ssdm data to ssdm portal####################################

        #check done status
        if statusDone[i]==1:
            print(schoolName[i] +" has already been synced")
            pass

        else:
            #read ssdm data from excel file
            df = pd.read_csv(fileNames[i])
            
            DateTime = df["Date & Time"].astype('string')

            day = []
            month = []
            year = []
            hour = []
            minute = []

            for date_string in DateTime:
                parts = date_string.split()  # Split the date and time parts
                date_part = parts[0].split('-')  # Split the date into year, month, and day
                time_part = parts[1].split(':')  # Split the time into hour, minute, and second

                year.append(date_part[0])
                month.append(date_part[1])
                day.append(date_part[2])
                hour.append(time_part[0])
                minute.append(time_part[1])

            student_name = df['Student']
            item = df['Item']
            teacher = df['Awarded by']
            synced = df['Synced'].to_numpy()

            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 10)
            driver.set_script_timeout(10)
            driver.get("https://ssdm.moe.gov.my/")

            login_box, pass_box = wait.until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR , "input[name='id'], input[name='password']")))
            login_box.send_keys(ssdmID[i])
            pass_box.send_keys(ssdmPass[i])
            wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@class='button']"))).click()

            totalStudent = len(student_name)
            print("Total submission: "+ str(totalStudent))

            for x in range(len(student_name)):
                #check if submission is already synced
                if synced[x]==1:
                    totalStudent-=1
                    print("This submission has already been synced")
                    print("Remaining submissions :"+ str(totalStudent))
                    continue
                name_search = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR , "input[name='nama']")))        
                name_search.send_keys(student_name[x])
                wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@name='cari']"))).click()
                fail_message = driver.find_elements(By.XPATH, "//span[text()='Maaf tiada padanan dengan data yang di cari dalam pangkalan data murid (APDM)']")
                #check if student is available in database, skip if not present
                if len(fail_message) >0:
                    print(student_name[x]+" is not available in the database ")
                    totalStudent-=1
                    print("Remaining submissions :"+ str(totalStudent))
                    #update sync status
                    synced[x]=1
                    df['Synced'] = synced
                    df.to_csv(fileNames[i], index=False)
                    continue
                wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'Papar Kes')]"))).click()
                wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Tambah Amalan Baik']"))).click()

                #fills out keterangan
                keterangan = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR , "textarea[name='keteranganamalanbaik']")))
                keterangan.send_keys(item[x])
                #select tempat
                select_tempat = driver.find_element(By.XPATH, "//select[@name='TEMPAT']")
                select = Select(select_tempat)
                select.select_by_value('01')
                #select amalan baik
                cat_ssdm = category_ssdm(item[x])

                #check if item is valid or not
                AB = dropdown_ssdm(cat_ssdm)
                if AB == "Invalid":
                    totalStudent-=1
                    print("Remaining submissions :"+ str(totalStudent))
                    driver.get("https://ssdm.moe.gov.my/kemaskini_murid.cfm")
                    continue

                #select date
                return_day(day[x])
                return_month(month[x])
                #select time
                return_hour(hour[x])
                return_min(minute[x])
                #select teacher
                select_cikgu = driver.find_element(By.XPATH, "//select[@name='papar_guru']")
                select = Select(select_cikgu)
                select.select_by_visible_text(teacher[x])
                #click checkbox
                driver.find_element(By.XPATH, "//input[@type='checkbox']").click()
                #submit ssdm
                driver.find_element(By.XPATH, "//input[@name='simpan_amalanbaik']").click()

                #update sync status
                synced[x]=1
                df['Synced'] = synced
                df.to_csv(fileNames[i], index=False)
                
                #handle alert popup
                alert = wait.until(EC.alert_is_present())
                if alert:
                    alert.accept()
                driver.switch_to.default_content()
                
                totalStudent-=1
                print("Sync is successful!")
                print("Remaining submissions :"+ str(totalStudent))

                driver.find_element(By.XPATH, "//a[@href='kemaskini_murid.cfm']").click()

            clear_output()
            print("Syncing for " +schoolName[i]+ " is done")

            #update school sync status
            statusDone[i]=1
            df_login['DONE'] = statusDone
            df_login.to_csv(loginCSV, index=False)

            driver.close()

    except Exception as e:
        print("An error occured:", e)
        sys.exit("Error encountered. Stopping the script.")  # Stop the script
    finally:
        close_driver()

print("Sync process is done for all schools")

