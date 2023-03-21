from selenium.webdriver.chrome.service import Service
#import undetected_chromedriver as uc
import pytest
from selenium.webdriver.support.wait import WebDriverWait
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By


service = Service(executable_path="/home/alex/PycharmProjects/pythonProject/chromedriver/chromedriver")
options = webdriver.ChromeOptions()
options.add_argument(f"user-agent='Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36'")
options.add_argument('--disable-blink-features=AutomationControlled')
#options.add_argument("--headless")

driver = webdriver.Chrome(service=service, options=options)


def is_visible_by_xpath(xpath):
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, xpath)))

def is_visible_by_ID(ID):
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, ID)))

def is_clickable_by_xpath(xpath):
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, xpath)))

#   присутствие
def is_presence_by_xpath(xpath):
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, xpath)))

def is_presence_by_ID(ID):
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, ID)))