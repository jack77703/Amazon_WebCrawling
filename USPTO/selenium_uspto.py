from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
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

def haswordmark(xpath):
    try:
        chrome.find_element_by_xpath(xpath)
        return True
    except:
        return False

def hasnextpage(xpath):
    try:
        chrome.find_element_by_xpath(xpath)
        return True
    except:
        return False

options = Options()
options.add_argument("--disable-notifications")
#用selenium操作瀏覽器並搜尋
chrome = webdriver.Chrome('/usr/local/bin/chromedriver', options=options)
chrome.get("http://tess2.uspto.gov/")
time.sleep(2)
chrome.find_element_by_xpath('/html/body/center/table[1]/tbody/tr[2]/td/font/font/a').click()
element = chrome.find_element_by_xpath('/html/body/form/font/table[4]/tbody/tr[1]/td/input')
chrome.find_element_by_xpath('/html/body/form/font/table[3]/tbody/tr[2]/td[2]/input').click()
time.sleep(2)
element.send_keys('taiwan')
time.sleep(2)
select = Select(chrome.find_element_by_xpath("/html/body/form/font/table[4]/tbody/tr[2]/td/select[@id='fieldtype']"))
select.select_by_visible_text(u"ALL")
chrome.find_element_by_xpath('/html/body/form/font/table[4]/tbody/tr[4]/td/input[3]').click()
search = chrome.find_element_by_xpath('/html/body/table[4]/tbody/tr/td[2]/input')
search.send_keys('8701')
chrome.find_element_by_xpath('/html/body/table[4]/tbody/tr/td[1]/form/input[3]').click()
#創excel檔
wb = openpyxl.Workbook()
sheet = wb.create_sheet("USPTO", 0)
titles = ("Word Name","Goods and Services","Mark Drawing Code","Owner")
sheet.append(titles)
i = 2
while(i < 52):
    if hasxpath('/html/body/table[7]/tbody/tr[%i]/td[2]/a'% (i,)) == True :
        chrome.find_element_by_xpath('/html/body/table[7]/tbody/tr[%i]/td[2]/a'% (i,)).click()
        if haswordmark("//table[5]//td/b[contains(text(),'Word Mark')]/../following-sibling::td") == False :
            i += 1
            chrome.back()
            if (i == 52): 
                if hasnextpage('//img[@alt="next TOC list"]') == True :
                    chrome.find_element_by_xpath('//img[@alt="next TOC list"]').click()
                    i = 2
                    pass
                else:
                    break
            continue
        Word_name = chrome.find_element_by_xpath("//table[5]//td/b[contains(text(),'Word Mark')]/../following-sibling::td").text
        Goods_and_Services = chrome.find_element_by_xpath("//table[5]//td/b[contains(text(),'Goods and Services')]/../following-sibling::td").text
        Mark_Drawing_Code = chrome.find_element_by_xpath("//table[5]//td/b[contains(text(),'Mark Drawing Code')]/../following-sibling::td").text
        Owner = chrome.find_element_by_xpath("//table[5]//td/b[contains(text(),'Owner')]/../following-sibling::td").text
        Info = (Word_name,Goods_and_Services,Mark_Drawing_Code,Owner)

        sheet.append(Info)
        chrome.back()
        time.sleep(1)
        i += 1
        if (i == 52):
            if hasnextpage('//img[@alt="next TOC list"]') == True :
                chrome.find_element_by_xpath('//img[@alt="next TOC list"]').click()
                i = 2
                pass
            else:
                break
    else:
            break

        

wb.save(os.path.expanduser("~/Desktop/USPTO/test4.xlsx"))



