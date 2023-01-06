from playwright.sync_api import sync_playwright
from scrapDocxs import dataWord
from utilities import pathsManager
import json
import os
rucNotFound=True
class scrapingPages:
    def __init__(self):
        self.urlSunat="https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/FrameCriterioBusquedaWeb.jsp"
        self.urlSunarp="https://enlinea.sunarp.gob.pe/sunarpweb/pages/acceso/ingreso.faces"
        self.elDni='https://eldni.com/pe/buscar-por-dni'
        self.userSunarp="leonelcg03"
        self.passSunarp="leonelcg03"
        self.browser=None
        self.context=None
        self.pageSunat=None
        self.pageSunarp=None
        self.pageDni=None
        self.x=pathsManager().currentFolderPath
    def gettingPagesData(self):
        with sync_playwright() as p:
            self.browser = p.chromium.launch(headless=False)
            self.context = self.browser.new_context()
            # starting pages ready to use a loop of search
            self.start_sunat()
            self.start_Sunarp()
            self.start_elDni()
            kardexsPath=os.path.join(self.x,"Kardexs")
            files=[file for file in os.listdir(kardexsPath) if file.endswith('.docx')]
            listDataWord=[]
            for file in files:
                dWord=dataWord(file)
                # llamar a mi funcion de busqueda de nombres naturales o juridicos
                listDataWord.append(dWord.getdataword())
            dataToConfront={}
            for kardexData in listDataWord:
                print(f"-------reading karex: {kardexData['kardex']}")
                dataToConfront[kardexData['kardex']]={}
                dniNames=[]
                dniName=None
                for dni in kardexData['dnis']:
                    dniName=self.get_dniName(dni)
                    dictNames={
                        "NRO DNI":dni,
                        "DNI ENCONTRADO":dniName,
                    }
                    dniNames.append(dictNames)
                if dniName!=None:
                    dataToConfront[kardexData['kardex']]['DNI']=dniNames

                rucsNames=[]
                rucName=None
                for ruc in kardexData['rucs']:
                    rucName=self.get_rucSunat(ruc)
                    dictRcus={
                        "NRO RUC":ruc,
                        "RUC ENCONTRADO":rucName,
                    }
                    rucsNames.append(dictRcus)
                if rucName!=None:
                    dataToConfront[kardexData['kardex']]['RUC']=rucsNames

                partids=[]
                partidName=None
                for partid in kardexData['partidaE']:
                    partidName=self.get_sunarpp(partid)
                    dicpartid={
                        "NRO PARTIDA":partid,
                        "PARTIDA ENCONTRADO":partidName
                    }
                    partids.append(dicpartid)
                if partidName!=None:
                    dataToConfront[kardexData['kardex']]['partidE']=partids
            PagesJsonPath=os.path.join(self.x,"dataofPages.json")
            with open(PagesJsonPath, 'w') as f:
                json.dump(dataToConfront, f,indent=4)
                    
    def start_sunat(self):
        self.pageSunat = self.context.new_page()
        self.pageSunat.goto(self.urlSunat)
    def get_rucSunat(self,ruc):
        rucNotFound=True
        while rucNotFound:
            try:
                self.pageSunat.query_selector("input#txtRuc").fill("")
                self.pageSunat.query_selector("input#txtRuc").fill(ruc)
                self.pageSunat.query_selector("button#btnAceptar").click(timeout=2000)
                self.pageSunat.wait_for_timeout(1000)
                self.pageSunat.wait_for_selector("div.col-sm-7 h4",timeout=2000)
                rucName=self.pageSunat.query_selector("div.col-sm-7 h4").inner_text()
                print(rucName)
                self.pageSunat.query_selector("button[class='btn btn-danger btnNuevaConsulta']").click()
                rucNotFound=False
            except:
                self.pageSunat.goto(self.urlSunat)
                self.pageSunat.wait_for_load_state(timeout=1000)
                rucName="Ruc not found"
                print("Ruc not found")
            return rucName
    def start_Sunarp(self):
        self.pageSunarp = self.context.new_page()
        self.pageSunarp.goto(self.urlSunarp)
        self.pageSunarp.query_selector("input[name='username']").fill(self.userSunarp)
        self.pageSunarp.query_selector("input[name='password']").fill(self.passSunarp)
        self.pageSunarp.query_selector("td.Ingresar").click()

    def get_sunarpp(self,partida):
        self.pageSunarp.wait_for_timeout(1000)
        frame=self.pageSunarp.frame(name='left_frame')
        frame.locator("//a[contains(text(),'Solicitar certificado literal de partida(copia literal)')]").click()
        try:
            self.pageSunarp.frame_locator("frame[name=\"main_frame\"]").frame_locator("frame[name=\"main_frame1\"]").locator("select[name=\"frmSolCertMenu\\:cboAreaRegistral\"]").select_option("22000")
            self.pageSunarp.frame_locator("frame[name=\"main_frame\"]").frame_locator("frame[name=\"main_frame1\"]").locator("select[name=\"frmSolCertMenu\\:cboCertificado\"]").select_option("70:1:L:Certificado Literal de Partida PJ:Propiedad Inmueble Predial")
            self.pageSunarp.frame_locator("frame[name=\"main_frame\"]").frame_locator("frame[name=\"main_frame1\"]").locator("button[role=\"button\"]:has-text(\"Solicitar\")").click()
            self.pageSunarp.frame_locator("frame[name=\"main_frame\"]").frame_locator("frame[name=\"main_frame1\"]").locator("select[name=\"frmPartidaDirecta\\:cmbOficina\"]").select_option("01|01")
            self.pageSunarp.frame_locator("frame[name=\"main_frame\"]").frame_locator("frame[name=\"main_frame1\"]").locator("[id=\"frmPartidaDirecta\\:radioPartida\\:1\"]").check()
            self.pageSunarp.frame_locator("frame[name=\"main_frame\"]").frame_locator("frame[name=\"main_frame1\"]").locator("input[role=\"textbox\"]").click()
            self.pageSunarp.frame_locator("frame[name=\"main_frame\"]").frame_locator("frame[name=\"main_frame1\"]").locator("input[role=\"textbox\"]").fill(partida)
            self.pageSunarp.frame_locator("frame[name=\"main_frame\"]").frame_locator("frame[name=\"main_frame1\"]").locator("button[role=\"button\"]:has-text(\"Buscar\")").click(timeout=6000)
            frame2=self.pageSunarp.frame(name="main_frame1")
            frame2.wait_for_selector("tbody tr[data-ri='0'] td:nth-child(6)",timeout=5000)
            direc=frame2.query_selector("tbody tr[data-ri='0'] td:nth-child(6)").inner_text()
            frame2.query_selector("button[id='frmResultadoPartDirecta:btnRegresar']").click()
            print(direc)
        except:
            direc="partida not found"
        return direc
    def start_elDni(self):
        self.pageDni = self.context.new_page()
        self.pageDni.goto(self.elDni)
        
    def get_dniName(self,dni):
        try:
            self.pageDni.query_selector("input#dni").fill("")
            self.pageDni.query_selector("input#dni").fill(dni)
            self.pageDni.query_selector("button[form='buscar-por-dni']").click()
            self.pageDni.wait_for_load_state()
            nombreCompleto=self.pageDni.query_selector("input#completos").get_attribute("value")
        except:
            nombreCompleto="dni not found"
        return nombreCompleto
        #self.pageDni.pause()
        
# scrpy=scrapingPages()
# scrpy.scrapInpages()