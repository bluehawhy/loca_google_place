from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import time
import chromedriver_autoinstaller

from _src import logger, excel

logging= logger.logger


def search_address(driver,Latitude=0,Longitude=0):
    url_google_plcae = 'https://www.google.com/maps/search/%s,+%s/data=!3m1!1e3?hl=ko' %(Latitude,Longitude)
    #logging.info(url_google_plcae)
    driver.get(url_google_plcae)
    wait = WebDriverWait(driver, 20)
    address_txt_xpath = '//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[10]/div/div[1]/span[3]/span[3]'
    address_plus_xpath = '//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[11]/div/div[1]/span[3]/span[3]'
    wait.until(EC.visibility_of_element_located((By.XPATH,address_txt_xpath)))
    current_address =driver.find_element("xpath",address_txt_xpath).text
    current_address_plus =driver.find_element("xpath",address_plus_xpath).text
    logging.info('address - %s address-plus - %s' %(current_address,current_address_plus))
    #time.sleep(1)

def selenium_start():
    #start selenium
    #set up chromedriver
    chrome_ver = chromedriver_autoinstaller.get_chrome_version().split('.')[0]  #크롬드라이버 버전 확인
    options = webdriver.ChromeOptions()
    options.add_argument('window-size=1920x1080')
    #options.add_argument('disable-gpu')
    options.add_argument('lang=ko_KR')
    options.add_argument('headless') # HeadlessChrome 사용시 브라우저를 켜지않고 크롤링할 수 있게 해줌
    #options.add_argument('User-Agent: Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.131 Safari/537.36')
    # 헤더에 headless chrome 임을 나타내는 내용을 진짜 컴퓨터처럼 바꿔줌.
    try:
        driver = webdriver.Chrome(f'./{chrome_ver}/chromedriver.exe',options=options)  
    except:
        chromedriver_autoinstaller.install(True)
        driver = webdriver.Chrome(f'./{chrome_ver}/chromedriver.exe',options=options)

    driver.implicitly_wait(10) #Implicitly wait for 10 sec
    return driver

driver = selenium_start()

file = r'C:/Users/miskang/Downloads/NG_List_E46.xlsx'

wb = excel.Workbook(file,read_only=False,data_only=False)
logwork_ws = wb.get_worksheet('NG_List')
logwork_row_index = wb.get_first_row('NG_List')
logging.info(logwork_row_index)

count = 0
for data in logwork_ws.rows:
    count += 1

    Latitude = data[3].value
    Longitude = data[2].value
    #logging.info('%s, %s' %(Latitude,Longitude))
    if count == 1:
        continue
    if count <= 3339:
        #break
        continue
    search_address(driver,Latitude,Longitude)

driver.close()
wb.close_workbook()