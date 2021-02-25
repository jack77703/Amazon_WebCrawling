from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
import time
import openpyxl
from bs4 import BeautifulSoup
import os

def hasxpath(xpath):
    try:
        chrome.find_element_by_xpath(xpath)
        return True
    except:
        return False

options = Options()
options.add_argument("--disable-notifications")
#用selenium操作瀏覽器並搜尋
chrome = webdriver.Chrome('/usr/local/bin/chromedriver', options=options)
chrome.get("https://www.taiwanexcellence.org/tw/award/product/result")
time.sleep(2)
#chrome.find_element_by_xpath('/html/body/center/table[1]/tbody/tr[2]/td/font/font/a').click()
#element = chrome.find_element_by_xpath('/html/body/form/font/table[4]/tbody/tr[1]/td/input')
#chrome.find_element_by_xpath('/html/body/form/font/table[3]/tbody/tr[2]/td[2]/input').click()
#time.sleep(2)
#element.send_keys('taiwan')
#time.sleep(2)
select = Select(chrome.find_element_by_xpath("/html/body/div[@class='wrap']/main[@class='award product result']/section[1]/div[@class='container']/div[@class='filter']/form/div[@class='col-xs-6 col-lg-2 offset-lg-2']/select[@class='custom-select']"))
select.select_by_visible_text(u"2016")
time.sleep(1)
chrome.find_element_by_xpath("/html/body/div[@class='wrap']/main[@class='award product result']/section[1]/div[@class='container']/div[@class='filter']/form/div[@class='col-xs-12 col-lg-3']/button").click()
#search = chrome.find_element_by_xpath('/html/body/table[4]/tbody/tr/td[2]/input')
#search.send_keys('8701')
#chrome.find_element_by_xpath('/html/body/table[4]/tbody/tr/td[1]/form/input[3]').click()
#創excel檔
wb = openpyxl.Workbook()
sheet = wb.create_sheet("Exllence", 0)
titles = ("Title","Type","Item Model","Company","Telephone","Fax","Email")
sheet.append(titles)
#next = 1
#for next in range(21):
#    chrome.find_element_by_xpath("//section/div/nav/ul/li/a[@title= 'Next']").click()
#    time.sleep(1)
#    next += 1

i = 1
while(i < 13):
    if hasxpath("//section/div/div[2]/a[%i]/div/div/img"% (i,)) == True :
        chrome.find_element_by_xpath("//section/div/div[2]/a[%i]/div/div/img"% (i,)).click()
        time.sleep(1)
        Title = chrome.find_element_by_xpath("//section/div/div/div/div/div[1]").text
        Type = chrome.find_element_by_xpath("//section/div/div/div/div/div[2]").text
        Item_Model = chrome.find_element_by_xpath("//section/div/div/div/div/div[3]").text
        Company = chrome.find_element_by_xpath("//section/div/div/div/div/div[4]").text
        chrome.find_element_by_xpath("//section/div/ul/li[4]/a").click()
        time.sleep(1)
        if hasxpath("/html/body/div[@class='wrap']/main[@class='award product detail']/section[@class='text-area container text-xs-center']/div[@class='row']/div[@class='tab-content text-xs-left col-xs-12 col-md-8 offset-md-2 ']/div[@id='contact']/p[2]/a[@title]")== True:
            Telephone = chrome.find_element_by_xpath("/html/body/div[@class='wrap']/main[@class='award product detail']/section[@class='text-area container text-xs-center']/div[@class='row']/div[@class='tab-content text-xs-left col-xs-12 col-md-8 offset-md-2 ']/div[@id='contact']/p[2]/a[@title]").text
        else:
            Telephone = "No Info"
        if hasxpath("/html/body/div[@class='wrap']/main[@class='award product detail']/section[@class='text-area container text-xs-center']/div[@class='row']/div[@class='tab-content text-xs-left col-xs-12 col-md-8 offset-md-2 ']/div[@id='contact']/p[3]/a[@title]") == True:
            Fax = chrome.find_element_by_xpath("/html/body/div[@class='wrap']/main[@class='award product detail']/section[@class='text-area container text-xs-center']/div[@class='row']/div[@class='tab-content text-xs-left col-xs-12 col-md-8 offset-md-2 ']/div[@id='contact']/p[3]/a[@title]").text
        else:
            Fax = "No Info"
        if hasxpath("/html/body/div[@class='wrap']/main[@class='award product detail']/section[@class='text-area container text-xs-center']/div[@class='row']/div[@class='tab-content text-xs-left col-xs-12 col-md-8 offset-md-2 ']/div[@id='contact']/p[4]/a[@title]")== True:
            Email = chrome.find_element_by_xpath("/html/body/div[@class='wrap']/main[@class='award product detail']/section[@class='text-area container text-xs-center']/div[@class='row']/div[@class='tab-content text-xs-left col-xs-12 col-md-8 offset-md-2 ']/div[@id='contact']/p[4]/a[@title]").text
        else:
            Email = "No Info"
        Info = (Title,Type,Item_Model,Company,Telephone,Fax,Email)
        sheet.append(Info)
        chrome.back()
        time.sleep(1)
        i += 1
        if(i == 13):
            chrome.find_element_by_xpath("//section/div/nav/ul/li/a[@title= 'Next']").click()
            i = 1
        wb.save(os.path.expanduser("~/Desktop/USPTO/2016.xlsx"))
    else:
        break
