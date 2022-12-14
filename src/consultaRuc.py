from playwright.sync_api import sync_playwright

url="https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/FrameCriterioBusquedaWeb.jsp"
url2="https://enlinea.sunarp.gob.pe/sunarpweb/pages/acceso/ingreso.faces"
rucNotFound=True
with sync_playwright() as p:
    global browser,context,page
    browser = p.chromium.launch(headless=False)
    context = browser.new_context()
            # Open new page
    page = context.new_page()
    while rucNotFound:
        try:
            page.goto(url)
            page.query_selector("input#txtRuc").fill("20256149681")
            page.query_selector("button#btnAceptar").click()
            page.wait_for_timeout(1000)
            page.wait_for_selector("div.col-sm-7 h4",timeout=2000)
            entire_name=page.query_selector("div.col-sm-7 h4").inner_text()
            rucNotFound=False
        except:
            print("Ruc not found")

    print(entire_name)
    page2 = context.new_page()
    page2.goto(url2)
    page2.query_selector("input[name='username']").fill("leonelcg03")
    page2.query_selector("input[name='password']").fill("leonelcg03")
    page2.query_selector("td.Ingresar").click()
    page2.wait_for_timeout(1000)
    #print(page2.url)
    frame=page2.frame(name='left_frame')
    frame.locator("//a[contains(text(),'Solicitar certificado literal de partida(copia literal)')]").click()
    page2.frame_locator("frame[name=\"main_frame\"]").frame_locator("frame[name=\"main_frame1\"]").locator("select[name=\"frmSolCertMenu\\:cboAreaRegistral\"]").select_option("22000")
    # Select 70:1:L:Certificado Literal de Partida PJ:Propiedad Inmueble Predial
    page2.frame_locator("frame[name=\"main_frame\"]").frame_locator("frame[name=\"main_frame1\"]").locator("select[name=\"frmSolCertMenu\\:cboCertificado\"]").select_option("70:1:L:Certificado Literal de Partida PJ:Propiedad Inmueble Predial")
    # Click button[role="button"]:has-text("Solicitar")
    page2.frame_locator("frame[name=\"main_frame\"]").frame_locator("frame[name=\"main_frame1\"]").locator("button[role=\"button\"]:has-text(\"Solicitar\")").click()
    # Select 01|01
    page2.frame_locator("frame[name=\"main_frame\"]").frame_locator("frame[name=\"main_frame1\"]").locator("select[name=\"frmPartidaDirecta\\:cmbOficina\"]").select_option("01|01")
    # Check [id="frmPartidaDirecta\:radioPartida\:1"]
    page2.frame_locator("frame[name=\"main_frame\"]").frame_locator("frame[name=\"main_frame1\"]").locator("[id=\"frmPartidaDirecta\\:radioPartida\\:1\"]").check()
    page2.frame_locator("frame[name=\"main_frame\"]").frame_locator("frame[name=\"main_frame1\"]").locator("input[role=\"textbox\"]").click()
    page2.frame_locator("frame[name=\"main_frame\"]").frame_locator("frame[name=\"main_frame1\"]").locator("input[role=\"textbox\"]").fill("00128244")
    page2.frame_locator("frame[name=\"main_frame\"]").frame_locator("frame[name=\"main_frame1\"]").locator("button[role=\"button\"]:has-text(\"Buscar\")").click()
    frame2=page2.frame(name="main_frame1")
    frame2.wait_for_selector("tbody tr[data-ri='0'] td:nth-child(6)")
    direc=frame2.query_selector("tbody tr[data-ri='0'] td:nth-child(6)").inner_text()
    print(direc)

    page3=context.new_page()
    page3.goto("https://eldni.com/")
    page3.query_selector("input#dni").fill("48023851")
    page3.query_selector("button[form='buscar-por-dni']").click()
    page3.wait_for_load_state()
    nombreCompleto=page3.query_selector("input#completos").get_attribute("value")
    print(nombreCompleto)
    page2.pause()
