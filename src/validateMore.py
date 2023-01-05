import docx
import re
from numerToLetters import numero_a_moneda_sunat
import os
from utilities import configData,pathsManager
def correlativos_correctos(correlativos):
  grupo_actual = correlativos[0]
  for i in range(len(correlativos) - 1):
    if correlativos[i] + 1 != correlativos[i + 1]:
      # Si el número actual no es el número anterior más uno,
      # verificamos si es el primer número de un nuevo grupo
      if correlativos[i + 1] == grupo_actual:
        grupo_actual = correlativos[i + 1]
      else:
        return False
  return True

def getText(doc):
    fullText=[]
    for para in doc.paragraphs:
        fullText.append(para.text)
    return " ".join(fullText)

def getAll_correlatives(doc):
    fullText=[]
    corr=[]
    for j,para in enumerate(doc.paragraphs):
        if para.text.find("Ultimo párrafo del documento antes del ultimo correlativo") != -1:
            print(para.text)
            print(doc.paragraphs[j+1].text)
            print(doc.paragraphs[j+2].text)
        # correlativ=re.findall(r"^([1-9]{1,2})",para.text)
        #correlativ2=re.findall(r"^[1-9]{1,2}\.\d{1,2}.",para.text)
        correlativ3=re.findall(r"^[1-9]{1,2}\.\d{1,2}\.(\d{1,2}).?",para.text)
        # correlativ4=re.findall(r"^[1-9]{1,2}\.\d{1,2}\.\d{1,2}\.(\d{1,2}).",para.text)
        # corr.extend(correlativ)
        #corr.extend(correlativ2)
        corr.extend(correlativ3)
        # corr.extend(correlativ4)
    # conver the corr list to int
    #print(corr)
    # corr = [int(i) for i in corr]
    # if len(corr) > 0:
    #     if correlativos_correctos(corr):
    #         print("CORRELATIVOS CORRECTOS")
    #     else:
    #         print("CORRELATIVOS INCORRECTOS")

#list of .docx files in Kardex folder

def correlativeLetters(doc):
    pm=pathsManager()
    cnfd=configData(os.path.join(pm.currentFolderPath,"config.xlsx"))
    numLeters=cnfd.get_data_config()
    numLeters=numLeters["CORRELATIVOS LETRAS"].values.tolist()
    #numLeters=["PRIMER","SEGUND","TERCER","CUART","QUINT","SEXT","SÉPTIM","OCTAV","NOVEN","DÉCIM"]
    pattern=""
    for i,leter in enumerate(numLeters):
        for j,c in enumerate(leter):
            if j==0:
                pattern0=r"^(%s *" % c
            else:
                pattern0=pattern0+r"%s *" % c
        if i==0:
            pattern=pattern0+r"[AO])"
        else:
            pattern=pattern+r"|%s[AO])" % pattern0
    setNumsList=[]
    for p in doc.paragraphs:
        setNums=re.findall(pattern,p.text)
        setNumsList.extend(setNums)

    netCorrelatives=[]
    for nums in setNumsList:
        for num in nums:
            if num:
                netCorrelatives.append(num)
    return netCorrelatives
def getIndexCorrelative(listOfCorrelatives):
    for i,cor in enumerate(listOfCorrelatives):
        if cor.replace(" ","").find("PRIMER") != -1:
            return i
def subMain():
    listOfDocxFiles=[f for f in os.listdir(r"C:\DanielBots\Sepulveda\sepulvedex\Kardexs") if f.endswith(".docx")]
    for file in listOfDocxFiles:
        print(file)
        doc=docx.Document(r"C:\DanielBots\Sepulveda\sepulvedex\Kardexs\\"+file)
        text=getText(doc)
        #getAll_correlatives(doc)
        correlativeLetters(doc)
    
subMain()
#listOfDocxFiles=[f for f in os.listdir(r"C:\DanielBots\Sepulveda\sepulvedex\Kardexs") if f.endswith(".docx")]

# for file in listOfDocxFiles:
#     print(file)
#     doc=docx.Document(r"C:\DanielBots\Sepulveda\sepulvedex\Kardexs\\"+file)
#     text=getalltext(doc)

