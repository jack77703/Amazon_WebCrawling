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
#select = Select(chrome.find_element_by_name(''))
#select.select_by_visible_text(u"ALL")
chrome.find_element_by_xpath('/html/body/form/font/table[4]/tbody/tr[4]/td/input[3]').click()
chrome.find_element_by_xpath('//img[@alt="next TOC list"]').click()


