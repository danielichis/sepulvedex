from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time
import highligth


def download():
    print("Inicio del Proceso")
    time.sleep(5)

    options = webdriver.ChromeOptions()

    options.add_argument("start-maximized")
    options.add_experimental_option('prefs', {
    "download.default_directory": "C:\\Datafiles", 
    "download.prompt_for_download": False, 
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True 
    })

    driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
    executor_url = driver.command_executor._url
    session_id = driver.session_id

    print (session_id)
    print (executor_url)
    print (driver.title)
    kardex_query = 'K45366'
    linkpage='http://192.168.0.90/legasys/www'
    driver.get(linkpage)
    print(driver.title)
    time.sleep(2)
    driver.find_element(By.XPATH,'/html/body/div[2]/form/div[1]/div[2]/input').click()
    driver.find_element(By.XPATH,'/html/body/div[2]/form/div[1]/div[2]/input').send_keys('BOT')
    time.sleep(1)
    driver.find_element(By.XPATH,'/html/body/div[2]/form/div[1]/div[3]/input').click()
    driver.find_element(By.XPATH,'/html/body/div[2]/form/div[1]/div[3]/input').send_keys('BOT')
    time.sleep(1)

    driver.find_element(By.XPATH,'/html/body/div[2]/form/div[2]/button[1]').click()
    time.sleep(2)

    driver.get('http://192.168.0.90/legasys/www/legal/consultas/consugeneral.php?mod=2')

    time.sleep(3)
    driver.find_element(By.XPATH,'/html/body/div[4]/div/table/tbody/tr[1]/td/form/table/tbody/tr[2]/td/div[1]/input[1]').click()
    driver.find_element(By.XPATH,'/html/body/div[4]/div/table/tbody/tr[1]/td/form/table/tbody/tr[2]/td/div[1]/input[1]').send_keys(kardex_query)


    time.sleep(1)
    #click en buscar Kardex
    driver.find_element(By.XPATH,'/html/body/div[4]/div/table/tbody/tr[1]/td/form/table/tbody/tr[2]/td/div[1]/input[2]').click()


    time.sleep(2)
    driver.find_element(By.XPATH,'/html/body/div[4]/div/table/tbody/tr[2]/td/div/center/div/div[3]/a/img').click()

    time.sleep(5)
    # storing the current window handle to get back to dashboard
    main_page = driver.current_window_handle
    print (main_page)
    # wait for page to load completely
    time.sleep(3)
    # changing the handles to access login page
    for handle in driver.window_handles:
        if handle != main_page:
            login_page = handle
            
    # change the control to signin page        
    driver.switch_to.window(login_page)

    link_escritura="/html/body/div/div/form/table/tbody/tr/td/table/tbody/tr[5]/td/table/tbody/tr[5]/td/ul/li[4]/a"
    driver.find_element(By.XPATH,link_escritura).click()
    time.sleep(3)

    driver.find_element(By.XPATH,'/html/body/div/div/form/table/tbody/tr/td/table/tbody/tr[5]/td/table/tbody/tr[5]/td/div/center/table/tbody/tr[2]/td[2]/img').click()
    print("Descarga Doc",kardex_query)                             

    time.sleep(2)


download()

highligth.writedoc()