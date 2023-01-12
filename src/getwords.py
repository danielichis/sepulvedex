from playwright.sync_api import sync_playwright
import pandas as pd
import os
from googleSheet import get_data_from_gsheet
from utilities import pathsManager
def download_legacy():
    user="BOT"
    password="BOT"
    url="http://192.168.0.90"
    url2="http://192.168.0.90/legasys/www/legal/consultas/consugeneral.php?mod=2"

    with sync_playwright() as p:
        global context,page,browser
        browser = p.chromium.launch(headless=False)
        context=browser.new_context()
        page = context.new_page()
        page.goto(url)
        page.query_selector("#usuario").fill(user)
        page.query_selector("#password").fill(password)
        page.query_selector("button[type='submit']").click()
        page.wait_for_timeout(3000)
        page.goto(url2)
        page.wait_for_selector("input[name='criterio']")
        kardexs=get_list_kardexs()
        for kardex in kardexs:
            try:
                get_kardex(kardex["kardex"])
            except Exception as e:
                print(e)
                kardexf=kardex["kardex"]
                print(f"descarga fallida {kardexf}")
                pass
        browser.close()
    return kardexs
def get_kardex(kardex):
    print(f"kardex: {kardex} descargando....")
    page.query_selector("input[name='criterio']").fill('')
    page.query_selector("input[name='criterio']").fill(kardex)
    page.query_selector("input[type='submit']").click()
    #new window
    with context.expect_page() as window:
        page.locator("//a[contains(text(),'Generar')]/img").click()
    page2 = window.value
    page2.wait_for_load_state()
    page2.locator("//a[contains(text(),'ESCRITURA')]").click()
    with page2.expect_download() as download_info:
        try:
            page2.wait_for_selector("img[src='/legasys/www/assets/images/genera_escritura.png']",timeout=1000)
            page2.query_selector("img[src='/legasys/www/assets/images/genera_escritura.png']").click()
            print("....")
        except Exception as e :
            print("....")
            page2.wait_for_selector("img[alt='Visualizar Documento']",timeout=1000)
            page2.query_selector("img[alt='Visualizar Documento']").click()
            #print(e)
        print("....")


    #new_window = window.value
    download = download_info.value
    page2.close()
    nameFile=kardex+".docx"
    pm=pathsManager().currentFolderPath
    nameFile=os.path.join(pm,"Kardexs",nameFile)
    download.save_as(nameFile)
    print("descargado")
    #driver.get('http://192.168.0.90/legasys/www/legal/consultas/consugeneral.php?mod=2')
    #page.pause()     

def get_list_kardexs():
    return get_data_from_gsheet()
#download_legacy()
#print(get_list_kardexs())