from playwright.sync_api import sync_playwright
import pandas as pd
import os
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
                get_kardex(kardex)
            except:
                pass
        browser.close()
def get_kardex(kardex):
    print(f"kardex: {kardex} descargando....")
    page.query_selector("input[name='criterio']").fill('')
    page.query_selector("input[name='criterio']").fill(kardex)
    page.query_selector("input[type='submit']").click()
    #new window
    with context.expect_page() as window:
        page.locator("//a[contains(text(),'Generar')]/img").click()
    page2 = window.value
    page2.locator("//a[contains(text(),'ESCRITURA')]").click()
    with page2.expect_download() as download_info:
        page2.wait_for_selector("div#ver_ver>a>img")
        page2.query_selector("div#ver_ver>a>img").click()
    #new_window = window.value
    download = download_info.value
    page2.close()
    nameFile=kardex+".docx"
    nameFile=os.path.join("Kardexs",nameFile)
    download.save_as(nameFile)

    #driver.get('http://192.168.0.90/legasys/www/legal/consultas/consugeneral.php?mod=2')
    #page.pause()     

def get_list_kardexs():
    fp=pathsManager().currentFolderPath
    path=os.path.join(fp,"CDCONF.xlsx")
    df=pd.read_excel(path)
    return df["k"].values.tolist()
#don_legacy()