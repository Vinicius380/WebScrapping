import time

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

service = Service() # e usada para instanciar o chrome webdriver
options = webdriver.ChromeOptions() # definir preferencias do chrome
driver = webdriver.Chrome(service=service, options=options)

url = "https://www.youtube.com/"

driver.get(url)
time.sleep(5)

driver.find_element(By.TAG_NAME, value="input").send_keys("Ghost")
driver.find_element(By.TAG_NAME, value="input").send_keys(Keys.ENTER)
time.sleep(1)
driver.find_element(By.LINK_TEXT, value="Ghost - Square Hammer (Official Music Video)").click()
time.sleep(5)
