#!/usr/bin/env python
# coding: utf-8

# In[4]:


from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import time
import shutil
import pandas as pd
from IPython.display import clear_output
import sys

login_excel = r"C:\Users\Aiman\Downloads\terengganu_school_detail.xlsx"
df = pd.read_excel(login_excel)
school = df['SCHOOL']
email = df['EMAIL']
done = df['DONE']

def close_driver():
    try:
        driver.quit()
    except Exception as e:
        print("An error occurred while closing the WebDriver:", e)

for y in range(len(school)):
    
    try:

        if done[y]==1:
            print(school[y] + " is already done")
            continue
        else:
            print("Downloading for " + school[y])

        #adjust path accordingly to Windpws/MacOS
        #directory = "/Users/piixel/Desktop/DATA SSDM PPD JB/SMK DATO USMAN AWANG"
        directory = "C:\\Users\\Aiman\\Desktop\\ssdm terengganu 1jun-31aug\\" + school[y]
        if os.path.exists(directory):
            shutil.rmtree(directory)  # Remove the existing directory and its contents
        os.makedirs(directory)

        options = webdriver.ChromeOptions()
        options.add_argument(f"--disable-gpu")
        options.add_experimental_option("prefs", {
            "download.default_directory": directory,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True
        })
        driver = webdriver.Chrome(options=options)

        password = 'password'
        x=0

        item = ['MEMBANTU GURU JIKA ADA MAJLIS',
                'MENOLONG DAN MENGAWAL KELAS SEMASA GURU TIADA',
                'MEMBANTU GURU MENGAMBIL KEHADIRAN PELAJAR',
                'MEMBANTU GURU JIKA MEMERLUKAN BANTUAN',
                'MEMBANTU MENGANGKAT BUKU/BARANG GURU',
                'DATANG AWAL KE SEKOLAH',
                'BERDISIPLIN',
                'BERCAKAP DENGAN SOPAN',
                'BERPAKAIAN KEMAS',
                'MENYAPA/MEMBERI SALAM KEPADA GURU',
                'MENYUSUN KEMAS MEJA',
                'MEMBUANG DEBU/HABUK DI CERMIN/MEJA/DINDING',
                'MEMBERSIHKAN PAPAN HITAM',
                'MENGEMOP BILIK',
                'MENYAPU BILIK-BILIK KHAS',
                'MENYUSUN MEJA/PERABOT',
                'MENGHIAS BILIK',
                'MEMBERSIHKAN MEJA/PERABOT YANG BERDEBU',
                'MOP LANTAI',
                'MENYAPU SAMPAH-SARAP DI SEKELILING KELAS',
                'MUDAH BERINTERAKSI ATAU BERGAUL',
                'MENGHORMATI PENDAPAT ORANG LAIN',
                'BIJAK BERKOMUNIKASI UNTUK MENDAPATKAN MAKLUMAT',
                'MENJADI FASILATATOR',
                'MELAPOR KES MASALAH SOSIAL YANG SERIUS',
                'MELAPOR KES MEMBAWA/MENGGUNAKAN BARANG TERLARANG',
                'MELAPOR KES TIDAK BERDISIPLIN',
                'MELAPOR KES LEWAT KE SEKOLAH',
                'MELAPOR KES PONTENG SEKOLAH/KELAS',
                'MENYUSUN DAN MENCANTIK KAN KAWASAN TAMAN/LANDSKAP',
                'MEMOTONG RUMPUT YANG PANJANG',
                'MEMBANTU MENYIRAM POKOK BUNGA',
                'MENGUTIP SAMPAH-SARAP YANG ADA',
                'MEMBANTU MENYAPU DAUN-DAUN KERING',
                'MELAKUKAN DENGAN PENUH SEMANGAT',
                'SUKA MENOLONG ORANG LAIN',
                'MEMBANTU MENGEMAS SEMULA BARANG YANG DIGUNAKAN',
                'BEKERJASAMA MENGECAT DINDING',
                'MENUNJUKKAN KREATIVITI']

        driver.get("https://admin.studentqr.com/login")
        print(driver.title)
        data=0

        #####################login process#############################
        wait = WebDriverWait(driver, 10, poll_frequency=0.1)
        login_box, pass_box = wait.until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR , "input[name='email'], input[name='password']")))
        login_box.send_keys(email[y])
        pass_box.send_keys(password)

        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']"))).click()
        #####################login process#############################

        #####################navigate to reporting####################
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/div[2]/div[1]/div[2]/img"))).click()
        time.sleep(2)
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/div[2]/div/div/div/div[1]/div/div[7]/a"))).click()
        #####################navigate to reporting####################

        #####################set filter###############################
        wait.until(EC.element_to_be_clickable((By.XPATH ,"//*[@id='secure']/div[2]/div[3]/div/div/div/div/div[2]/div/div[1]/button/span[1]"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH ,"/html/body/div/div[1]/div[2]/div[3]/div/div/div/div/div[2]/div/div[2]/div/a"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='List']"))).click()

        wait.until(EC.element_to_be_clickable((By.XPATH ,"//*[@id='secure']/div[2]/div[3]/div/div/div/div/div[3]/div/button"))).click()
        for i in item:
            try:
                wait.until(EC.element_to_be_clickable((By.XPATH ,"//span[text()='"+i+"']"))).click()
            except:
                continue

        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@class='button is-success']"))).click()

        wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Today']"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/div[2]/div[3]/div/div/div/div/div[4]/div/div[2]/div/a[3]"))).click()
        #####################set filter###############################

        #####################set date range#############################
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/div[2]/div[3]/div/div/div/div/div[5]/div/div/div[1]/button"))).click()
        while x<3:
            wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='app']/div[2]/div/div/div[1]/button[1]"))).click()
            time.sleep(0.5)
            x=x+1
        wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@data-date='2023-06-01']"))).click()
        x=0
        while x<2:
            wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='app']/div[2]/div/div/div[1]/button[2]"))).click()
            time.sleep(0.5)
            x=x+1
        wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@data-date='2023-08-31']"))).click()
        #####################set date range#############################

        #####################generate data##############################
        wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='secure']/div[2]/div[3]/div/div/div/div/div[7]/button"))).click()
        try:
            element = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div[1]/div[2]/div[4]/div/div/div/div[3]/div/div[3]/div[3]/div/div/div[1]/button/span[1]")))
            element.click()
        except:
            #update sync status
            df.loc[y, 'DONE'] = 1
            df.to_excel(login_excel, index=False)

            clear_output()
            print(school[y] + " does not have any submissions")
            time.sleep(1)
            driver.close()
            continue
        #wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div[1]/div[2]/div[4]/div/div/div/div[3]/div/div[3]/div[3]/div/div/div[1]/button/span[1]"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div[1]/div[2]/div[4]/div/div/div/div[3]/div/div[3]/div[3]/div/div/div[2]/div/a[4]"))).click()
        #####################generate data##############################

        #####################select all and download###################
        count=2
        wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='secure']/div[2]/div[4]/div/div/div/div[3]/div/div[1]/div[1]/div/button/input"))).click()#SELECT ALL
        wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='secure']/div[2]/div[4]/div/div/div/div[3]/div/div[1]/div[1]/div/div/div[1]/button/span[1]"))).click()#ACTION
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div[2]/div[4]/div/div/div/div[3]/div/div[1]/div[1]/div/div/div[2]/div/a[2]"))).click()#DOWNLOAD
        time.sleep(2)

        try:
            #next button
            wait.until(EC.visibility_of_element_located((By.XPATH, "//button[contains(text(),'{}')]".format(count))))
        except:
            #if submission is less than 100, close the sessions
            #update sync status
            df.loc[y, 'DONE'] = 1
            df.to_excel(login_excel, index=False)

            clear_output()
            print(school[y] + " is done")
            time.sleep(1.5)
            driver.close()
            continue

        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'{}')]".format(count)))).click()
        data=data+100
        print('printed :' +str(data))
        print('page :' +str(count-1))
        count=count+1

        while(True):
            try:
                #next page
                wait.until(EC.visibility_of_element_located((By.XPATH, "//button[contains(text(),'{}')]".format(count))))
            except:
                break
            #select button for pages other than page 1 need to be double clicked, for some reason
            wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='secure']/div[2]/div[4]/div/div/div/div[3]/div/div[1]/div[1]/div/button/input"))).click()
            wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='secure']/div[2]/div[4]/div/div/div/div[3]/div/div[1]/div[1]/div/button/input"))).click()
            wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='secure']/div[2]/div[4]/div/div/div/div[3]/div/div[1]/div[1]/div/div/div[1]/button/span[1]"))).click()
            wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div[2]/div[4]/div/div/div/div[3]/div/div[1]/div[1]/div/div/div[2]/div/a[2]"))).click()
            time.sleep(1.5)
            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'{}')]".format(count)))).click()
            data=data+100
            print('printed :' +str(data))
            print('page :' +str(count-1))
            count=count+1


        #for last page loop
        wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='secure']/div[2]/div[4]/div/div/div/div[3]/div/div[1]/div[1]/div/button/input"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='secure']/div[2]/div[4]/div/div/div/div[3]/div/div[1]/div[1]/div/button/input"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='secure']/div[2]/div[4]/div/div/div/div[3]/div/div[1]/div[1]/div/div/div[1]/button/span[1]"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div[2]/div[4]/div/div/div/div[3]/div/div[1]/div[1]/div/div/div[2]/div/a[2]"))).click()
        time.sleep(1.5)
        data=data+100
        print('printed :' +str(data))
        print('page :' +str(count-1))
        count=count+1

        #update sync status
        df.loc[y, 'DONE'] = 1
        df.to_excel(login_excel, index=False)

        clear_output()
        print(school[y] + " is done")
        driver.close()
    
    except Exception as e:
        print("An error occured:", e)
        sys.exit("Error encountered. Stopping the script.")  # Stop the script
    finally:
        close_driver()

print("All school is done")


# In[ ]:




