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


# chrome.find_element_by_xpath('/html/body/center/table[1]/tbody/tr[2]/td/font/font/a').click()
#element = chrome.find_element_by_xpath('/html/body/form/font/table[4]/tbody/tr[1]/td/input')
# chrome.find_element_by_xpath('/html/body/form/font/table[3]/tbody/tr[2]/td[2]/input').click()
# time.sleep(2)
# element.send_keys('taiwan')
# time.sleep(2)
#select = Select(chrome.find_element_by_xpath("/html/body/div[@class='wrap']/main[@class='award product result']/section[1]/div[@class='container']/div[@class='filter']/form/div[@class='col-xs-6 col-lg-2 offset-lg-2']/select[@class='custom-select']"))
# select.select_by_visible_text(u"2016")
# time.sleep(1)
#chrome.find_element_by_xpath("/html/body/div[@class='wrap']/main[@class='award product result']/section[1]/div[@class='container']/div[@class='filter']/form/div[@class='col-xs-12 col-lg-3']/button").click()
#search = chrome.find_element_by_xpath('/html/body/table[4]/tbody/tr/td[2]/input')
# search.send_keys('8701')
# chrome.find_element_by_xpath('/html/body/table[4]/tbody/tr/td[1]/form/input[3]').click()
# 創excel檔
wb = openpyxl.Workbook()
sheet = wb.create_sheet("Taipei", 0)
titles = ("Company", "Telephone", "Address", "Main_Product","Introduction", "Email", "Website")
sheet.append(titles)
#next = 1
# for next in range(21):
#    chrome.find_element_by_xpath("//section/div/nav/ul/li/a[@title= 'Next']").click()
#    time.sleep(1)
#    next += 1


links = [link.get_attribute('href') for link in chrome.find_elements_by_xpath("//div[@class='group_cat']/table[@class='category']/tbody/tr/td[@class='link_title']/a")]
for link in links:
    chrome.get(link)
    if hasxpath("//div[@class='company']") == True:
        Company = chrome.find_element_by_xpath("//div[@class = 'company']").text
    else:
        Company = "No Info"
    if hasxpath("//div[@class='row nspList'][3]/div[@class='fleft value']/span[@class='sub'][2]") == True:
        Telephone = chrome.find_element_by_xpath("//div[@class='row nspList'][3]/div[@class='fleft value']/span[@class='sub'][2]").text
    else:
        Telephone = "No Info"
    if hasxpath("//div[@class='fleft value']/div[@class='mb10']") == True:
        Address = chrome.find_element_by_xpath("//div[@class='fleft value']/div[@class='mb10']").text
    else:
        Address = "No Info"
    if hasxpath("//div[@class='mr_b20'][1]") == True:
        Main_Product = chrome.find_element_by_xpath("//div[@class='mr_b20'][1]").text
    else:
        Main_Product = "No Info"
    if hasxpath("//div[@class='mr_b20'][2]") == True:
        Introduction = chrome.find_element_by_xpath("//div[@class='mr_b20'][2]").text
    else:
        Introduction = "No Info"
    if hasxpath("//div[@class='row nspList'][7]/div[@class='fleft value']") == True:
        Email = chrome.find_element_by_xpath("//div[@class='row nspList'][7]/div[@class='fleft value']").text
    else:
        Email = "No Info"
    if hasxpath("//div[@class='row nspList'][8]/div[@class='fleft value']") == True:
        Website = chrome.find_element_by_xpath("//div[@class='row nspList'][8]/div[@class='fleft value']").text
    else:
        Website = "No Info"

    Info = (Company, Telephone, Address, Main_Product,Introduction, Email, Website)
    sheet.append(Info)
    time.sleep(1)
    wb.save(os.path.expanduser("~/Desktop/Amazon Crawling/Industrial Accociation/Taipei.xlsx"))
    chrome.back()
    