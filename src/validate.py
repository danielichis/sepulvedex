from docx.enum.text import WD_COLOR_INDEX
from docx import Document
from utilities import pathsManager
import re
import os
import json
#Instalarte ademas este modulo para los comments
#pip install bayoo-docx

def split_text(text, word):
    pattern = re.compile(r'([\S\s]*)(\b{})([\S\s]*)'.format(word))
    #pattern = re.compile(r'[\S\s]*(\d{8})[\S\s]*'.format(re.escape(word)))
    match = pattern.search(text)
    if match:
        return match.groups()
    return None

def highlight_wordRun(run,paragraph,numero,descripction,comment):
    texRun=run.text
    before, word, after = split_text(texRun, numero)
    run.text = ""
    run.add_text(before)
    #run.add_break()
    if comment:
        run=paragraph.add_run()
        run.add_comment(f'{descripction} No se encuentra en el documento',author='BOT CONFRONT')
    run=paragraph.add_run()
    run.add_text(word)
    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    #run.add_text(word).font.highlight_color = WD_COLOR_INDEX.YELLOW
    #paragraph.add_comment('Otra Observacion Encontrada',author='Observacion de DNI')
    #paragraph.runs[2].font.highlight_color = WD_COLOR_INDEX.YELLOW
    #run.add_comment('Otra Observacion Encontrada',author='Observacion de DNI')
    run=paragraph.add_run()
    run.add_text(after)
    return paragraph
def highlight_word(pathkard,doc, numero,descripction,mutliple):
    alltext=getText(doc)
    print(f"reading {numero} {descripction}")
    if isPresent(descripction,alltext):
        for paragraph in doc.paragraphs:
            runes=paragraph.runs
            for i,run in enumerate(runes):
                if run.text.find(numero)>-1:
                    paragraph=highlight_wordRun(run,paragraph,numero,descripction,False)
        doc.save(pathkard)
        return
    else:
        print(f"no se encontro {descripction}")
        pass
    print(f"cantidad de parrafos {len(doc.paragraphs)}")
    for paragraph in doc.paragraphs:
        print(f"reading parrafo {paragraph.text}")
        if paragraph.text.find(numero) == -1:
            pass
        else:
            runes=paragraph.runs
            replaced=False
            for i,run in enumerate(runes):
                if run.text.find(numero) == -1:
                    pass
                else:
                    paragraph=highlight_wordRun(run,paragraph,numero,descripction,True)
                    replaced=True
                if mutliple==False and replaced==True:
                    break
            if mutliple==False and replaced==True:
                break
    doc.save(pathkard)
    #print("Documento Verificado")
def getText(doc):
    fullText=[]
    for para in doc.paragraphs:
        fullText.append(para.text)
    return " ".join(fullText)
def isPresent(Dname,alltext):
    if alltext.find(Dname) == -1 :
        return False
    else:
        return True
def validateAllData(value,doc,pathkard):
    alltext=getText(doc)
    for key2 in value:
        #print(f"reading ------------{key2}")
        value2=value[key2]
        for dict3 in value2:
            values=list(dict3.values())
            #print(f"reading-----{values}")
            highlight_word(pathkard,doc, values[0],values[1],False)

            
    
def readJsonPages():
    pm=pathsManager()
    jsonPath=os.path.join(pm.currentFolderPath,"dataofPages.json")
    with open(jsonPath) as json_file:
        data = json.load(json_file)    
    for key in data:
        print(f"reading ------------{key}----------")
        pathkard=os.path.join(pm.currentFolderPath,"Kardexs",key)
        pathkardOut=os.path.join(pm.currentFolderPath,"KardexsOut",key)
        #print(pathkard)
        doc=Document(pathkard)
        value = data[key]
        validateAllData(value,doc,pathkardOut)


readJsonPages()
#doc=Document(r"C:\DanielBots\Sepulveda\sepulvedex\src\daniel.docx")
#highlight_word(r"C:\DanielBots\Sepulveda\sepulvedex\src\daniel2.docx",doc,"8888","Danielito",False)
#highlight_word(r"C:\DanielBots\Sepulveda\sepulvedex\src\daniel.docx",'párrafo',False)


# text = 'Hola, esto es una prueba de texto y esto es otra prueba de texto y esto es la ultima prueba de texto'

# before, word, after = split_text(text, 'prueba')
# print('Antes:', before)
# print('Palabra:', word)
# print('Después:', after)