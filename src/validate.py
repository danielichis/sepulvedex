from docx.enum.text import WD_COLOR_INDEX
from docx import Document
from utilities import pathsManager
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE
from numerToLetters import numero_a_moneda_sunat
import re
import os
import json
#Instalarte ademas este modulo para los comments
#pip install bayoo-docx

def split_text(text, word):
    pattern = re.compile(r'([\S\s]*)(\b{})([\S\s]*)'.format(word))
    match = pattern.search(text)
    if match:
        return match.groups()
    return None

def split_Runs(doc,word):
    for p in doc.paragraphs:
        if p.text.find(word) != -1:
            virtualRuns=p.runs
            p.text = ""
            for r in virtualRuns:
                if r.text.find(word) != -1:
                    before, word, after = split_text(r.text, word)
                    #clone the format of the original run
                    p.add_run(before,style='CommentsStyle')
                    # runBefore.font.name = fontName
                    # runBefore.font.size = fontSize
                    #p.add_run(before)
                    commentRun=p.add_run("",style='CommentsStyle')
                    # commentRun.font.name = fontName
                    # commentRun.font.size = fontSize
                    wordRun=p.add_run(word,style='CommentsStyle')
                    # wordRun.font.name = fontName
                    # wordRun.font.size = fontSize
                    affterRun=p.add_run(after,style='CommentsStyle')
                    # affterRun.font.name = fontName
                    # affterRun.font.size = fontSize
                else:
                    p.add_run(r.text,style='CommentsStyle')
    return doc
def style_Token(doc,word,comment):
    for p in doc.paragraphs:
        for i,r in enumerate(p.runs):
            if p.runs[i].text.find(word) != -1:
                p.runs[i].font.highlight_color = WD_COLOR_INDEX.YELLOW
                if comment:
                    p.runs[i-1].add_comment(f'{word} No se encuentra en el documento',author='BOT CONFRONT')
                    #r.add_comment(f'{word} No se encuentra en el documento',author='BOT CONFRONT')
    return doc

def confront(pathkard,doc, numero,descripction,splitMode):
    alltext=getText(doc)
    print(f"reading {numero} {descripction}")
    if splitMode:
        doc=split_Runs(doc,numero)
    else:
        if isPresent(descripction,alltext):
            doc=style_Token(doc,numero,False)
        else:
            doc=style_Token(doc,numero,True)
    doc.save(pathkard)
def getText(doc):
    fullText=[]
    for para in doc.paragraphs:
        fullText.append(para.text)
    return " ".join(fullText)
def getNumbers(doc):
    fullText=[]
    numerLetters=[]
    for para in doc.paragraphs:
        fullText.append(para.text)
        numletter=re.findall(r"S\/ (\d{1,}.*?\.\d{2}) \((.*?)\)",para.text)
        if len(numletter)>0:
            numerLetters.append(numletter[0])
    return numerLetters
def isPresent(Dname,alltext):
    if alltext.find(Dname) == -1 :
        return False
    else:
        return True
def validateAllData(value,doc,pathkard,splitMode):
    for key2 in value:
        value2=value[key2]
        for dict3 in value2:
            values=list(dict3.values())
            confront(pathkard,doc, values[0],values[1],splitMode)

def readJsonPages():
    pm=pathsManager()
    jsonPath=os.path.join(pm.currentFolderPath,"dataofPages.json")
    with open(jsonPath) as json_file:
        data = json.load(json_file)    
    for key in data:
        print(f"reading ------------{key}----------")
        pathkard=os.path.join(pm.currentFolderPath,"Kardexs",key)
        pathkardOut=os.path.join(pm.currentFolderPath,"KardexsOut",key)
        if key=="K42218.docx":
            #print(pathkard)
            doc=Document(pathkard)
            font_styles = doc.styles
            paragraphs_styles=[s for s in font_styles if s.type == WD_STYLE_TYPE.PARAGRAPH]
            paragraphs_styles_names=[styles.type for styles in paragraphs_styles]
            font_charstyle = font_styles.add_style('CommentsStyle', WD_STYLE_TYPE.CHARACTER)
            help(font_styles.add_style)
            font_object = font_charstyle.font
            font_object.size = Pt(10)
            font_object.name = 'Arial Narrow'
            value = data[key]
            validateAllData(value,doc,pathkardOut,True)
            validateAllData(value,doc,pathkardOut,False)

readJsonPages()