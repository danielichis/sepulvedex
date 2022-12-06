#import playwright
from playwright.sync_api import sync_playwright 

def found_name(dni):
    page.query_selector("input[class='form-input']").fill(dni)
    page.query_selector("button[type='submit']").click()
    #page.pause()
    try:
        page.wait_for_selector("input#completos")
        page.set_default_timeout(2000)
        nombreCompleto=page.query_selector("input#completos").get_attribute("value")
    except:
        nombreCompleto="No encontrado"
    #page.pause()
    page.query_selector("input[class='form-input']").fill('')
    page.screenshot(path="example.png")
    
    print(nombreCompleto)

listaDnis=["48023851","70027986","19982788","19811114","23150497","19979598"]
# for dni in listaDnis:
#     initpage()
#     found_name(dni)
with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()
    page.goto("https://eldni.com/pe/buscar-por-dni")
    #page.pause()
    for dni in listaDnis:
        print(f"buscando dni: {dni}")
        found_name(dni)
    browser.close()
#get_name("73710270")