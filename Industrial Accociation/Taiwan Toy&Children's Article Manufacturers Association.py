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
chrome.get("https://www.tcma.com.tw/company.php?page_num=3&type=all")
time.sleep(1)

# chrome.find_element_by_xpath('/html/body/center/table[1]/tbody/tr[2]/td/font/font/a').click()
#element = chrome.find_element_by_xpath('/html/body/form/font/table[4]/tbody/tr[1]/td/input')
# chrome.find_element_by_xpath('/html/body/form/font/table[3]/tbody/tr[2]/td[2]/input').click()
# time.sleep(2)
# element.send_keys('taiwan')
# time.sleep(2)
#select = Select(chrome.find_element_by_xpath("//strong/form[@id='form']/select[@id='jumpMenu']"))
# select.select_by_visible_text(u"62")
# time.sleep(1)
#chrome.find_element_by_xpath("/html/body/div[@class='wrap']/main[@class='award product result']/section[1]/div[@class='container']/div[@class='filter']/form/div[@class='col-xs-12 col-lg-3']/button").click()
#search = chrome.find_element_by_xpath('/html/body/table[4]/tbody/tr/td[2]/input')
# search.send_keys('8701')
# chrome.find_element_by_xpath('/html/body/table[4]/tbody/tr/td[1]/form/input[3]').click()
# 創excel檔
wb = openpyxl.Workbook()
sheet = wb.create_sheet("Taiwan Toy", 0)
titles = ("Company", "Telephone", "Address", "Email", "Website")
sheet.append(titles)
#next = 1
# for next in range(21):
#    chrome.find_element_by_xpath("//section/div/nav/ul/li/a[@title= 'Next']").click()
#    time.sleep(1)
#    next += 1

i = 1
while(i < 28):
    if hasxpath("//li[@class='col-3 link_hover'][%i]/div[@class='Tit']/a" % (i,)) == True:
        chrome.find_element_by_xpath("//li[@class='col-3 link_hover'][%i]/div[@class='Tit']/a" % (i,)).click()
        time.sleep(1)
        if hasxpath("//div[@class='product_Right_title']") == True:
            Company = chrome.find_element_by_xpath("//div[@class='product_Right_title']").text
        else:
            Company = "No Info"
        if hasxpath("//div[@class='product_Right_content'][2]") == True:
            Telephone = chrome.find_element_by_xpath("//div[@class='product_Right_content'][2]").text
        else:
            Telephone = "No Info"
        if hasxpath("//div[@class='product_Right_content'][4]") == True:
            Address = chrome.find_element_by_xpath("//div[@class='product_Right_content'][4]").text
        else:
            Address = "No Info"
        if hasxpath("//div[@class='product_Right_content'][5]") == True:
            Email = chrome.find_element_by_xpath("//div[@class='product_Right_content'][5]").text
        else:
            Email = "No Info"
        if hasxpath("//div[@class='product_Right_content'][6]") == True:
            Website = chrome.find_element_by_xpath("//div[@class='product_Right_content'][6]").text
        else:
            Website = "No Info"

        Info = (Company, Telephone, Address, Email, Website)
        sheet.append(Info)
        chrome.back()
        time.sleep(1)
        i += 1
        wb.save(os.path.expanduser("~/Desktop/Amazon Crawling/Industrial Accociation/Taiwan_Toy_3.xlsx"))
        if(i == 28):
            chrome.find_element_by_xpath("//tr/td[5]/a[@class='pagelink_no']").click()
            i = 1
    else:
        break
