import json
import win32com.client
import os
from pathlib import Path
import docx2txt
import sys

class highlighting:

    def __init__(self):
        self.words = []


    def highlightingKardex(self, words, kardex):
        word = win32com.client.Dispatch("Word.Application")
        documents = 'Documents'
        # kardexFolder = os.path.join(documents, 'Kardex')
        # kardexDoc = kardex + '.docx'
        kardexPath = os.path.join('Kardexs', kardex)
        doc = word.Documents.Open(kardexPath)

        #paragraph = doc.paragraphs(0)
        
        for paragraph in doc.paragraphs:
            # print('-----------------------------------------------------------------------------------------------------------------------------------------------------')
            data = paragraph.Range.Text
            
            for i in words:
                try:
                    word.Selection.Find.Execute(i)
                    #win32com.client.constants.wdYellow
                    word.Selection.Range.HighlightColorIndex = 7
                
                except:
                    pass
        
        
        resultingKardex = os.path.join(Path(getCurrentPath()).parent, 'resultingKardex')
        if os.path.exists(resultingKardex) == False:
            os.mkdir(resultingKardex)
        
        l = len(kardex)
        l = l - 5
        newName = kardex[:l] + '-2.docx'
        resultingName = os.path.join(resultingKardex, newName)
        doc.SaveAs(resultingName)
        doc.Close()


    def highlightingEveryKardex(self, jsonName):
        jsonPath = os.path.join(Path(getCurrentPath()).parent, jsonName)
        with open(jsonPath, 'r') as f:
            data = json.load(f)

            for i in list(data.keys()):
                kardexPath = os.path.join(Path(getCurrentPath()).parent, 'Kardexs')
                kardexPath = os.path.join(kardexPath, i)
                
                text = docx2txt.process(kardexPath)
                for j in list(data[i].keys()):
                    for k in data[i][j]:
                        idEncontrado = j + ' ENCONTRADO'
                        nroID = 'NRO ' + j
                        if k[idEncontrado] == 'not found':
                            self.words.append(k[nroID])

                        if k[idEncontrado] in text == False:
                            self.words.append(k[nroID])
                self.highlightingKardex(self.words, i)
            

def getCurrentPath():   
    config_name = 'myapp.cfg'
    # determine if application is a script file or frozen exe
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    elif __file__:
        application_path = os.path.dirname(__file__)
    application_path2 = Path(application_path)
    return application_path2.absolute()

    
if __name__ == '__main__':
    x = highlighting()
    x.highlightingEveryKardex('dataofPages.json')
    # print(Path(getCurrentPath()).parent)