from selenium import webdriver
import time
import tkinter

chromewebdriver = "C:/Arbeitsplatz/python/raspiprojekt/chromedriver"
browser = webdriver.Chrome(executable_path= chromewebdriver)
browser.implicitly_wait(0.5)
'browser.maximize_window()'
browser.get("https://solarbotics.com/product/600559/")
'browser.fullscreen_window()'
testElement = browser.find_element_by_id('menu-item-1021')
testElement.click()
time.sleep(1)
testElement = browser.find_element_by_id("menu-item-119314")
testElement.click()
time.sleep(2)
testElement.back()
testElement = browser.find_element_by_id('menu-item-1021')
time.sleep(1)
testElement.click()
'''testElement = browser.find_element_by_name("s")
testElement.send_keys("seleiunm")
testElement.send_keys( u'\ue007')
PATH=https://stackoverrun.com/de/q/10559530 Taste drueken ganz unten auf der Seite
PATH= https://selenium-python.readthedocs.io/api.html#selenium.webdriver.common.keys.Keys.RETURN '''

testElement = browser.find_element_by_id("menu-item-119324")
testElement.click()
time.sleep(3)
testElement = browser.find_element_by_id('menu-item-84')
testElement.click()
time.sleep(3)
testElement.back()