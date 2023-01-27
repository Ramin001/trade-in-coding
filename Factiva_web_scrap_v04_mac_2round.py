
"""
Created on Wed Jul  4 17:11:53 2018

@author: seungho.back16
"""
# %reset -f

#directory setting

import os
#os.getcwd()
#os.chdir('G:/My Drive/02. PhD/01. Courses/07. 2018 Summer/02. 2nd Year Paper/01. Python code/1. Result')


#Import all you need 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select


import random
import time
import re
import sys
import pandas as pd
from pandas import ExcelWriter
import numpy as np

#chrome driver and factiva login

#driver = webdriver.Chrome('C:/Downloads/chromedriver')
driver = webdriver.Chrome('/Users/Andy/Downloads/chromedriver')
wait = WebDriverWait(driver, 1000)
url = "http://myaccess.library.utoronto.ca/login?url=https://global.factiva.com/en/sess/login.asp?xsid=S003Gzo3Wvf5DEs5DEmNDAsNpQtOTZyMHmnRsIuMcNG1pRRQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQQAA"
driver.get(url)

'''
#Use below if you need to log in

driver.find_element_by_xpath('//*[@id="visualWrapper"]/div[2]/p/strong[1]/a').click()
id = driver.find_element_by_id('inputID')
id.send_keys("backseun")

id = driver.find_element_by_id('inputPassword')
id.send_keys("xxxxx")
driver.find_element_by_xpath('//*[@id="query"]/button').click()
'''

#inputs (interval, number of loops)
t1 = 10
t2 = 20
l1 = 93
l2 = 119

#set up the source as wire
wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'div.pill')))
driver.find_element_by_css_selector('div.pill').click()
driver.find_element_by_css_selector('div.remove').click()
driver.find_element_by_css_selector('div.pnlTabArrow').click()
wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#scMnu > ul > li:nth-child(23) > a.mnuItm')))
driver.find_element_by_css_selector('#scMnu > ul > li:nth-child(23) > a.mnuItm').click() 
driver.find_element_by_css_selector('div.pnlTabArrow').click()


'''
#search keywords only in the headline
driver.find_element_by_css_selector('#moreOptsBtn > div.pnlTabArrow').click() 
select = Select(driver.find_element_by_css_selector('#sfd'))
select.select_by_visible_text('Headline')
'''

'''
#open the spreadsheet and define the dimension of your spreadsheet
mydata = pd.read_excel('/Users/Andy/Downloads/Factiva Search Result_Round4.xlsx')
nrow = len(mydata.index)
ncol = len(mydata.columns)
'''

'''
#make a search result matrix
matrix = np.zeros((nrow,11))
sr = pd.DataFrame(matrix,columns=['alliance ID','Seller','Buyer','date_year','date_month','Type','stage','subject','disease','technology','search result' ])
mylist = [(0,0),(1,4),(2,5),(3,6),(4,7),(5,10),(6,11),(7,12),(8,24),(9,26)]
'''

#look for search area and clear 
search = driver.find_element_by_name("ftx")
search.clear()

#search starts
## keyword setting
for index in range (l1,l2):
    search_key1 = mydata.iloc[int(index),4]
    search_key2 = mydata.iloc[int(index),5]
    search_keys="%s and %s" %(search_key1,search_key2)
    alliance_year = mydata.iloc[int(index),6]
    alliance_month = mydata.iloc[int(index),7]
    start_date=[1,1,int(alliance_year)]
    end_date=[12,31,int(alliance_year)]
    search = driver.find_element_by_name("ftx")
    search.send_keys(search_keys)

## dates setting
    if start_date == None:
        driver.find_element_by_css_selector("select#dr > option[value='_Unspecified']").click()
    else:
        driver.find_element_by_css_selector("#dr > option:nth-child(10)").click()
                                            
    driver.find_element_by_css_selector("#frd").clear()
    driver.find_element_by_css_selector("#frm").clear()
    driver.find_element_by_css_selector("#fry").clear()
    driver.find_element_by_css_selector("#frd").send_keys(start_date[1])
    driver.find_element_by_css_selector("#frm").send_keys(start_date[0])
    driver.find_element_by_css_selector("#fry").send_keys(start_date[2])
    
    
    driver.find_element_by_css_selector("#tod").clear()
    driver.find_element_by_css_selector("#tom").clear()
    driver.find_element_by_css_selector("#toy").clear()
    driver.find_element_by_css_selector("#tod").send_keys(end_date[1])
    driver.find_element_by_css_selector("#tom").send_keys(end_date[0])
    driver.find_element_by_css_selector("#toy").send_keys(end_date[2])
                                        
    driver.find_element_by_id("btnSBSearch").click()         
    
#search ends and download rtf and clear results    
    time.sleep (8)

#when there is a search result, download rtf
    if driver.find_elements_by_css_selector('#selectAll > input:nth-child(1)'):
        sr.iloc[int(index),10]=1
        driver.find_element_by_css_selector("#selectAll > input:nth-child(1)").click()
        driver.find_element_by_css_selector("#listMenuRoot > li.ppsrtf.ppsscroll.ppsscrollvisible > a").click()
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#listMenu-id-3 > li:nth-child(3) > a")))
        driver.find_element_by_css_selector("#listMenu-id-3 > li:nth-child(3) > a").click()
       
                                    
#I add this "try, except" because of possible errors here, in case error, close and reopen 
      
        try:
            element = driver.find_element_by_css_selector('body > h1')
            driver.close()
            driver = webdriver.Chrome('/Users/Andy/Downloads/chromedriver')
            #driver = webdriver.Chrome('C:/Downloads/chromedriver')
            wait = WebDriverWait(driver, 1000)
            url = "http://myaccess.library.utoronto.ca/login?url=https://global.factiva.com/en/sess/login.asp?xsid=S003Gzo3Wvf5DEs5DEmNDAsNpQtOTZyMHmnRsIuMcNG1pRRQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQQAA"
            driver.get(url)
            
            '''
            #Use below if you need to log in

            driver.find_element_by_xpath('//*[@id="visualWrapper"]/div[2]/p/strong[1]/a').click()
            id = driver.find_element_by_id('inputID')
            id.send_keys("backseun")
            
            id = driver.find_element_by_id('inputPassword')
            id.send_keys("Qortmdgh!11")
            driver.find_element_by_xpath('//*[@id="query"]/button').click()
            '''
            
            #set up the source as wire
            wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'div.pill')))
            driver.find_element_by_css_selector('div.pill').click()
            driver.find_element_by_css_selector('div.remove').click()
            driver.find_element_by_css_selector('div.pnlTabArrow').click()
            wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#scMnu > ul > li:nth-child(22) > a.mnuItm')))
            driver.find_element_by_css_selector('#scMnu > ul > li:nth-child(23) > a.mnuItm').click() 
            driver.find_element_by_css_selector('div.pnlTabArrow').click()
            
            '''
            #search keywords only in the headline
            driver.find_element_by_css_selector('#moreOptsBtn > div.pnlTabArrow').click() 
            select = Select(driver.find_element_by_css_selector('#sfd'))
            select.select_by_visible_text('Headline')
            '''
            
            #look for search area and clear 
            search = driver.find_element_by_name("ftx")
            search.clear()
            search.send_keys(search_keys)
            if start_date == None:
                driver.find_element_by_css_selector("select#dr > option[value='_Unspecified']").click()
            else:
                driver.find_element_by_css_selector("#dr > option:nth-child(10)").click()
            driver.find_element_by_css_selector("#frd").clear()
            driver.find_element_by_css_selector("#frm").clear()
            driver.find_element_by_css_selector("#fry").clear()
            driver.find_element_by_css_selector("#frd").send_keys(start_date[1])
            driver.find_element_by_css_selector("#frm").send_keys(start_date[0])
            driver.find_element_by_css_selector("#fry").send_keys(start_date[2])
            
            
            driver.find_element_by_css_selector("#tod").clear()
            driver.find_element_by_css_selector("#tom").clear()
            driver.find_element_by_css_selector("#toy").clear()
            driver.find_element_by_css_selector("#tod").send_keys(end_date[1])
            driver.find_element_by_css_selector("#tom").send_keys(end_date[0])
            driver.find_element_by_css_selector("#toy").send_keys(end_date[2])
                                                
            driver.find_element_by_id("btnSBSearch").click()         
            
            #search ends and download rtf and clear results    
            time.sleep (8)
            
            if driver.find_elements_by_css_selector('#selectAll > input:nth-child(1)'):
                sr.iloc[int(index),10]=1
        
                driver.find_element_by_css_selector("#selectAll > input:nth-child(1)").click()
                driver.find_element_by_css_selector("#listMenuRoot > li.ppsrtf.ppsscroll.ppsscrollvisible > a").click()
                wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#listMenu-id-3 > li:nth-child(3) > a")))
                driver.find_element_by_css_selector("#listMenu-id-3 > li:nth-child(3) > a").click()
                if driver.find_elements_by_css_selector('#maxWordsForPpsNotification > div > div > ul > li:nth-child(2) > div > span'): 
                    driver.find_element_by_css_selector("#maxWordsForPpsNotification > div > div > ul > li:nth-child(2) > div > span").click()
                    time.sleep(20)   
                else:
                    time.sleep(0.5)
                  
                wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#btnModifySearch > div > span")))
                driver.find_element_by_css_selector("#btnModifySearch > div > span").click()
                wait.until(EC.visibility_of_element_located((By.NAME, 'ftx')))
                search = driver.find_element_by_name("ftx")
                search.clear()   
                
            else:
                sr.iloc[int(index),10]=0
                wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#btnModifySearch > div > span")))
                driver.find_element_by_css_selector("#btnModifySearch > div > span").click()
                wait.until(EC.visibility_of_element_located((By.NAME, 'ftx')))
                search = driver.find_element_by_name("ftx")
                search.clear()                                     
            
        except:
            wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#btnModifySearch > div > span")))
            driver.find_element_by_css_selector("#btnModifySearch > div > span").click()
            wait.until(EC.visibility_of_element_located((By.NAME, 'ftx')))
            search = driver.find_element_by_name("ftx")
            search.clear()   
               
    
    else:
        sr.iloc[int(index),10]=0
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#btnModifySearch > div > span")))
        driver.find_element_by_css_selector("#btnModifySearch > div > span").click()
        wait.until(EC.visibility_of_element_located((By.NAME, 'ftx')))
        search = driver.find_element_by_name("ftx")
        search.clear()   

#save search result    
    for (x,y) in mylist:
        sr.iloc[int(index),int(x)]=mydata.iloc[int(index),int(y)]


#after the loop, save the search result as xlsx file
print ("I finished looking at press releases up to row",int(l2-1))
writer = ExcelWriter ('ROUND4 alliances result.xlsx')
sr.to_excel(writer,'search result'  )
writer.save()   

