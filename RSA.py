from conftest import driver, is_clickable_by_xpath, is_presence_by_xpath, is_presence_by_ID
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from time import sleep
from selenium.webdriver.support import expected_conditions as EC
import openpyxl
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
url = 'https://oto-register.autoins.ru/pto'

try:

    driver.get(url=url)
    driver.maximize_window()
    workbook = openpyxl.Workbook()
    worksheet = workbook.active


    # scroll to desired page
    for i in range(1, 2):
        next_page = driver.find_element(By.XPATH, '//*[@class="page-item next"]')
        driver.execute_script("arguments[0].scrollIntoView();", next_page)
        if is_clickable_by_xpath('//*[@class="page-item next"]'):
            pass
        else:
            print()
        next_page.click()
        sleep(1)
    for i in range(1, 131):
        elements = driver.find_elements(By.XPATH, '//*[@class="table_row"]')
        count = 0
        sleep(3)

        for element in elements:
            count +=1
            new_element = driver.find_element(By.XPATH, f'(//*[@class="table_row"])[{count}]')
            #driver.execute_script("arguments[0].scrollIntoView();", new_element)

            is_clickable_by_xpath(f'(//*[@class="table_row"])[{count}]')
            pto_id = driver.find_element(By.XPATH, f'(//*[@class="table_row"])[{count}]').get_attribute("data-request-id")
            categories = driver.find_element(By.XPATH, f"(//*[@class='table_row'])[{count}]/td[5]").text
            element.click()

            is_presence_by_xpath(f"//*[@class='red-part otoNum']")

            oto_number = driver.find_element(By.XPATH, '//*[@class="red-part otoNum"]').text
            try:
                status = driver.find_element(By.XPATH, '//*[@class="status-block ok"]/p').text
            except NoSuchElementException:
                status = driver.find_element(By.XPATH, '//*[@class="status-block pause"]/p').text



            addres = driver.find_element(By.XPATH, "//div[h4/text()='Адрес']/p").text
            limit_pto = driver.find_element(By.XPATH,"//div[h4/text()='Пропускная способность']/p").text
            phone = driver.find_element(By.XPATH, "//div[h4/text()='Телефон']/p").text
            mail = driver.find_element(By.XPATH, "//div[h4/text()='E-mail']/a").get_attribute("href")
            mail = mail.replace("mailto:", '')
            full_name = driver.find_element(By.XPATH, "//div[@class='leftPanel']/div[3]/p").text



            is_clickable_by_xpath('//*[@class="icon i-close"]')
            close = driver.find_element(By.XPATH, '//*[@class="close-popup"]').click()
            sleep(1)

            driver.execute_script("window.open('https://oto-register.autoins.ru/pto','_blank');")

            # переключаемся на новую вкладку
            driver.switch_to.window(driver.window_handles[1])
            is_clickable_by_xpath('(//*[@class="item"])[1]')
            PDL = driver.find_element(By.XPATH, '(//*[@class="item"])[1]').click()
            is_presence_by_ID("otoId")
            search = driver.find_element(By.ID, "otoId")
            search.click()
            search.send_keys(oto_number)
            search.send_keys(Keys.ENTER)
            is_presence_by_xpath(f"//*[@class='tac']")
            sleep(2)
            try:
                pdl_num = driver.find_element(By.XPATH, '(//*[@class="tac"])[1]').text
                if pdl_num == oto_number:
                    pdl = "есть"
                else:
                    pdl = "нет"

            except NoSuchElementException:
                pdl = "нет"

            is_clickable_by_xpath('(//*[@class="item"])[1]')
            PDL = driver.find_element(By.XPATH, '(//*[@class="item"])[2]').click()
            is_presence_by_ID("otoId")
            search = driver.find_element(By.ID, "otoId")
            search.click()
            search.send_keys(oto_number)
            search.send_keys(Keys.ENTER)
            is_presence_by_xpath("//*[@class='table_row']")
            try:

                driver.find_element(By.XPATH, "//*[@class='table_row']").click()
                sleep(1)
                try:
                    status_att = driver.find_element(By.XPATH, '//div[@class="status-block ok"]/p').text
                    if status_att == "Аттестат действителен":
                        stat_att = status_att
                except NoSuchElementException:
                    status_att = driver.find_element(By.XPATH, '//div[@class="status-block pause"]/p').text
                    if status_att == "Аттестат приостановлен":
                        stat_att = status_att


            except NoSuchElementException:
                driver.find_element(By.ID, "showCanceled1").click()
                driver.find_element(By.XPATH, '//*[@class="btn blue-btn"]').click()
                is_presence_by_xpath("//*[@class='table_row']")
                try:
                    driver.find_element(By.XPATH, "//*[@class='table_row']").click()


                except Exception as ex:
                    print(ex)

            print( oto_number,  "|" ,  status,"|" ,full_name, "|" , addres, "|" , limit_pto, "|" , phone, "|" , mail, "|" ,categories, "|", pto_id,"|", pdl ,"|", stat_att)


            row = [oto_number, status, full_name, addres, limit_pto, phone, mail, categories, pto_id, pdl, stat_att]
            worksheet.append(row)

            driver.close()
            driver.switch_to.window(driver.window_handles[0])
            continue

        next_page = driver.find_element(By.XPATH, '(//*[@class="page-link"])[13]').click()
        sleep(2)

except Exception as ex:
    print(ex)
    workbook.save('RSA.xlsx')

finally:
    workbook.save('RSA.xlsx')
    driver.close()
    driver.quit()
