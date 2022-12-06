from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time




def valida_partidas():

      print("Inicio del Proceso")
      # Primera Parte Lee Excel

      #time.sleep(15)

      options = webdriver.ChromeOptions()

      #options = Options()
      options.add_argument("start-maximized")
      options.add_experimental_option('prefs', {
      "download.default_directory": "C:\\Datafiles", 
      "download.prompt_for_download": False, 
      "download.directory_upgrade": True,
      "plugins.always_open_pdf_externally": True 
      })
      #s = Service('C:\\BrowserDrivers\\chromedriver.exe')
      #driver = webdriver.Chrome(service=s, options=options)
      #driver.get("https://www.axisbank.com/interest-rate-on-deposits?cta=homepage-rhs-fd")
      #WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//h2[text()='Domestic Fixed Deposits']//following-sibling::div[1]/span"))).click()



      driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)

      #driver = webdriver.Chrome(r'chromedriver.exe')
      executor_url = driver.command_executor._url
      session_id = driver.session_id

   
      print (session_id)
      print (executor_url)
      print (driver.title)

      driver.get("https://enlinea.sunarp.gob.pe/sunarpweb/pages/acceso/ingreso.faces")
      print(driver.title)


      time.sleep(1)
      driver.find_element(By.XPATH,'//*[@id="contenedor"]/form/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td[2]/input').click()
      driver.find_element(By.XPATH,'//*[@id="contenedor"]/form/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td[2]/input').send_keys('leonelcg03')
      time.sleep(1)
      driver.find_element(By.XPATH,'//*[@id="contenedor"]/form/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[3]/td[2]/input[1]').click()
      driver.find_element(By.XPATH,'//*[@id="contenedor"]/form/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[3]/td[2]/input[1]').send_keys('leonelcg03')


      driver.find_element(By.XPATH,'//*[@id="contenedor"]/form/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[4]/td/a').click()
      
      time.sleep(3)
      

      #driver.execute_script("window.close('https://enlinea.sunarp.gob.pe/sunarpweb/pages/acceso/menuInformacion.html');");
      
      
      
      driver.execute_script("window.open('https://enlinea.sunarp.gob.pe/sunarpweb/pages/publicidadCertificada/menuCertificadoLiteral.faces','_blank');");
      
      #driver.find_element(By.XPATH,'/html/body/table/tbody/tr/td/form/div/div[2]/table/tbody/tr[1]/td[2]/select').click()
      
      

      
      #driver.close ( );

      time.sleep(25)


      #AQUI VA 2CAPTCHA CONTRATAR SERVICIO 1000 consultas por dolar
      #driver.find_element_by_xpath('/html/body/app-root/app-ingreso/body/div/div/div/div/div/form/div[5]/div/table/tr/td[3]/input').send_keys('3266517')


      # YA DENTRO DE SIGUELO
      # Partidas Vinculadas
      #driver.find_element_by_xpath('/html/body/app-root/app-titulo/body/div[5]/div[2]/div/textarea').click()
      #driver.find_element_by_xpath('/html/body/app-root/app-titulo/body/div[5]/div[2]/div/textarea').send_keys('1299417712994177')


      # check status
      #x = driver.find_element_by_xpath('//*[@id="RESULT_TextField-1"]').is_displayed()
      #print(x)
      #.get_attribute("value")
if 1==2:
   ha_presentacion = driver.find_element(By.XPATH,'/html/body/app-root/app-titulo/body/div[6]/div[2]/div/input').get_attribute("value")
   fecha_vencimiento = driver.find_element(By.XPATH,'/html/body/app-root/app-titulo/body/div[6]/div[3]/div/input').get_attribute("value")
   partidas_vinculadas = driver.find_element(By.XPATH,'/html/body/app-root/app-titulo/body/div[5]/div[2]/div/textarea').get_attribute("value")

   estado_titulo = driver.find_element(By.XPATH,'/html/body/app-root/app-titulo/body/div[8]/div[2]/div/div[1]').text


   print("fecha de presentacion:",fecha_presentacion)
   print("fecha de vencimiento:",fecha_vencimiento)
   print("Partidas Vinculadas:",partidas_vinculadas)
   print("Estado del titulo:",estado_titulo)

   time.sleep(10)

   if estado_titulo=="INSCRITO":
      driver.find_element(By.XPATH,'/html/body/app-root/app-titulo/body/div[11]/div[3]/table/tr[2]/td/a/span').click()
      time.sleep(8)
      driver.find_element(By.XPATH,'/html/body/div[2]/div[2]/div/mat-dialog-container/app-partidas/mat-card-content/div/div/table/tbody/tr/td[3]/button').click()
      print("Proceso Terminado...")
      time.sleep(15)
      



   writeexcel()
   print("Muestro Excel con Resultados")
   time.sleep(35)

   #driver.close()

