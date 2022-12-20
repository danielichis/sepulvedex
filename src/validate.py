from docx.enum.text import WD_COLOR_INDEX
from docx import Document
from utilities import pathsManager
import re
import os
import json
#Instalarte ademas este modulo para los comments
#pip install bayoo-docx

def highlight_word(path, word,mutliple=False):
    doc=Document(path)
    for paragraph in doc.paragraphs:
        if paragraph.text.find(word) == -1:
            pass
        else:
            print(paragraph.text)
            texParragraph=paragraph.text
            pattern=r"(.*)%s(.*)" % word
            tbf=re.findall(pattern,texParragraph)[0][0]
            tba=re.findall(pattern,texParragraph)[0][1]
            paragraph.text = ""
            paragraph.add_run(tbf)
            paragraph.add_run().add_comment('Otra Observacion Encontrada',author='Observacion de DNI')
            paragraph.add_run(word).font.highlight_color = WD_COLOR_INDEX.YELLOW
            #paragraph.add_comment('Otra Observacion Encontrada',author='Observacion de DNI')
            #paragraph.runs[2].font.highlight_color = WD_COLOR_INDEX.YELLOW
            #run.add_comment('Otra Observacion Encontrada',author='Observacion de DNI')
            paragraph.add_run(tba)
            if mutliple==False:
                break
    doc.save('daniel2.docx')
    print("Documento Verificado")
def getText(doc):
    fullText=[]
    for para in doc.paragraphs:
        fullText.append(para.text)
    return " ".join(fullText)
def isPresent(value,alltext):
    if alltext.find(value) == -1 :
        return False
    else:
        return True
def validateAllData(value,doc):
    alltext=getText(doc)
    for key2 in value:
        print(f"reading ------------{key2}")
        value2=value[key2]
        for dict3 in value2:
            print(f"reading-----{dict3}")
    
def readJsonPages():
    pm=pathsManager()
    jsonPath=os.path.join(pm.currentFolderPath,"dataofPages.json")
    with open(jsonPath) as json_file:
        data = json.load(json_file)    
    for key in data:
        print(f"reading ------------{key}----------")
        pathkard=os.path.join(pm.currentFolderPath,"Kardexs",key)
        print(pathkard)
        doc=Document(pathkard)
        value = data[key]
        validateAllData(value,doc)
readJsonPages()
#highlight_word(r"C:\DanielBots\Sepulveda\sepulvedex\src\daniel.docx",'párrafo',False)