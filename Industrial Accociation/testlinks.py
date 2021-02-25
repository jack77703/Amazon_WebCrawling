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
chrome.get("http://www.iatc.org.tw/memlist.html")
time.sleep(2)


links = [link.get_attribute('href') for link in chrome.find_elements_by_xpath("//div[@class='group_cat']/table[@class='category']/tbody/tr/td[@class='link_title']/a")]
for link in links:
    chrome.get(link)
    time.sleep(1)
    chrome.back()