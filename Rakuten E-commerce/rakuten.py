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
# 用selenium操作瀏覽器並搜尋
chrome = webdriver.Chrome('/usr/local/bin/chromedriver', options=options)
chrome.get("https://search.rakuten.co.jp/search/mall/台湾製/100026/?&v=2")
time.sleep(3)

wb = openpyxl.Workbook()
sheet = wb.create_sheet("rakuten", 0)
titles = ("title")
sheet.append([titles])

i = 1
while(i < 46):
    if hasxpath("//div[@class='dui-item searchresultitem'][%i]/div[@class='content']/div[@class='description title']/h2/a"% (i,))== True:
        title = chrome.find_element_by_xpath(("//div[@class='dui-item searchresultitem'][%i]/div[@class='content']/div[@class='description title']/h2/a")% (i,)).text
    else:
        title = "No Info"

    sheet.append([title])
    i += 1
    wb.save(os.path.expanduser("~/Desktop/Amazon Crawling/rakuten_new.xlsx"))
    if(i == 45):
        if hasxpath("//div[@class='dui-container pagination _centered']/div[@class='dui-pagination']/a[@class='item -next nextPage']")== True:
            chrome.find_element_by_xpath("//div[@class='dui-container pagination _centered']/div[@class='dui-pagination']/a[@class='item -next nextPage']").click()
            i = 1
        else:
            break
