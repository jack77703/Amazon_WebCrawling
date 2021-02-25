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
chrome.get("http://www.tccia.org.tw/member.php?proid=20160523110509&prokid=20160506085059")
time.sleep(1)

#chrome.find_element_by_xpath('/html/body/center/table[1]/tbody/tr[2]/td/font/font/a').click()
#element = chrome.find_element_by_xpath('/html/body/form/font/table[4]/tbody/tr[1]/td/input')
#chrome.find_element_by_xpath('/html/body/form/font/table[3]/tbody/tr[2]/td[2]/input').click()
#time.sleep(2)
#element.send_keys('taiwan')
#time.sleep(2)
select = Select(chrome.find_element_by_xpath("//strong/form[@id='form']/select[@id='jumpMenu']"))
select.select_by_visible_text(u"62")
#time.sleep(1)
#chrome.find_element_by_xpath("/html/body/div[@class='wrap']/main[@class='award product result']/section[1]/div[@class='container']/div[@class='filter']/form/div[@class='col-xs-12 col-lg-3']/button").click()
#search = chrome.find_element_by_xpath('/html/body/table[4]/tbody/tr/td[2]/input')
#search.send_keys('8701')
#chrome.find_element_by_xpath('/html/body/table[4]/tbody/tr/td[1]/form/input[3]').click()
#創excel檔
wb = openpyxl.Workbook()
sheet = wb.create_sheet("Exllence", 0)
titles = ("Company","Type","Item","Address","Telephone")
sheet.append(titles)
#next = 1
#for next in range(21):
#    chrome.find_element_by_xpath("//section/div/nav/ul/li/a[@title= 'Next']").click()
#    time.sleep(1)
#    next += 1

i = 2
while(i < 22):
    if hasxpath("//tbody/tr[%i]/td[@class='style8']/a"% (i,)) == True :
        chrome.find_element_by_xpath("//tbody/tr[%i]/td[@class='style8']/a"% (i,)).click()
        time.sleep(1)
        Company = chrome.find_element_by_xpath("//tbody/tr[1]/td[@class='style8']").text
        Type = chrome.find_element_by_xpath("//tbody/tr[2]/td[@class='style8']").text
        Item = chrome.find_element_by_xpath("//tbody/tr[3]/td[@class='style8']").text
        Address = chrome.find_element_by_xpath("//tbody/tr[4]/td[@class='style8']").text
        Telephone = chrome.find_element_by_xpath("//tbody/tr[5]/td[@class='style8'][2]").text
        if hasxpath("//tbody/tr[1]/td[@class='style8']")== True:
            Company = chrome.find_element_by_xpath("//tbody/tr[1]/td[@class='style8']").text
        else:
            Company = "No Info"
        if hasxpath("//tbody/tr[2]/td[@class='style8']") == True:
            Type = chrome.find_element_by_xpath("//tbody/tr[2]/td[@class='style8']").text
        else:
            Type = "No Info"
        if hasxpath("//tbody/tr[3]/td[@class='style8']")== True:
            Item = chrome.find_element_by_xpath("//tbody/tr[3]/td[@class='style8']").text
        else:
            Item = "No Info"
        if hasxpath("//tbody/tr[4]/td[@class='style8']")== True:
            Address = chrome.find_element_by_xpath("//tbody/tr[4]/td[@class='style8']").text
        else:
            Address = "No Info"
        if hasxpath("//tbody/tr[5]/td[@class='style8'][2]")== True:
            Telephone = chrome.find_element_by_xpath("//tbody/tr[5]/td[@class='style8'][2]").text
        else:
            Telephone = "No Info"
        Info = (Company,Type,Item,Address,Telephone)
        sheet.append(Info)
        chrome.back()
        i += 1
        wb.save(os.path.expanduser("~/Desktop/Amazon Crawling/Industrial Accociation/Hsinchu_6.xlsx"))
        if(i == 21):
            chrome.find_element_by_xpath("//tbody/tr/td/div/a[@class='page-right']").click()
            i = 2
    else:
        break
