from playwright.sync_api import sync_playwright
import pytesseract
import base64
from PIL import Image
from io import BytesIO
from base64 import b64decode
with sync_playwright() as p:
    global browser,context,page
    browser = p.chromium.launch(headless=False)
    context = browser.new_context()
            # Open new page
    page = context.new_page()
    page.goto("https://sistemasdgc.rree.gob.pe/carext_consulta_webapp")
    captcha=''
    k=0
    while captcha=='':
        if k>0:
            page.query_selector("#lbnActualizarCaptcha").click()
        src=page.query_selector("#imgCaptcha").get_attribute("src")
        imagestr = src
        im = Image.open(BytesIO(b64decode(imagestr.split(',')[1])))
        im.save("captcha.png")
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        captcha=pytesseract.image_to_string("captcha.png")
        k+=1
    page.query_selector("input[name='txtCodCaptcha']").fill(captcha)
    page.query_selector("input[name='txtNumeroConsulta']").fill("004680394")
    page.query_selector("a#btnBuscar").click()
    
    #save img
    #img.save_as("src\captcha.png")
    #page.pause()

