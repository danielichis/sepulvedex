import docx
from utilities import configData
from utilities import pathsManager
import os
import re
#get current path
class dataWord:
    def __init__(self,path) -> None:
        pm=pathsManager()
        self.kardexsPath=os.path.join(pm.currentPath,"Kardexs",path)
        self.configPath=os.path.join(pm.currentPath,"config.xlsx")
        self.doc=docx.Document(self.kardexsPath)
        self.fullText=self.getText()
        self.configData=configData(self.configPath)
        
        
        self.dnis=[]
        self.rucs=[]
        self.partidaE=[]
        self.carnets=[]

    def getText(self):
        fullText=[]
        for para in self.doc.paragraphs:
            r=para.add_run()
            fullText.append(para.text)
        self.fullText=" ".join(fullText)

    def get_names(self,n,nameKeys):
        dnislist=[]
        #self.fullText="DOCUMENTO NACIONAL DE IDENTIDAD NUMERO 12345678"
        for nameKey in nameKeys:
            if nameKey!="":
                reguexStr='(\d{%s})' % n
                pattern = r'%s %s' % (nameKey, reguexStr)
                dnis=re.findall(pattern,self.fullText)
                if dnis!=[]:
                    for d in dnis:
                        dnislist.append(d)
        return dnislist
    def getdataword(self):
        dniKeys=self.configData.df["NRO DNI"].values.tolist()
        rucKeys=self.configData.df["RUC"].values.tolist()
        ceKeys=self.configData.df["CE"].values.tolist()
        partidaKeys=self.configData.df["PARTIDA"].values.tolist()
        nationalityKeys=self.configData.df["NACIONALIDAD"].values.tolist()
        sombrearKeys=self.configData.df["SOMBREAR"].values.tolist()
        correlativeKeys=self.configData.df["CORRELATIVOS"].values.tolist()
        self.getText()
        dnis=self.get_names(8,dniKeys)
        dnis=list(dict.fromkeys(dnis))
        rucs=self.get_names(11,rucKeys)
        rucs=list(dict.fromkeys(rucs))
        partidaE=self.get_names(8,partidaKeys)
        partidaE=list(dict.fromkeys(partidaE))
        carnets=self.get_names(9,ceKeys)
        carnets=list(dict.fromkeys(carnets))
        #remove duplicates in list
        dataWord={
            "dnis":dnis,
            "rucs":rucs,
            "partidaE":partidaE,
            "carnets":carnets
        }
        return dataWord
def sub_main():
    files=[file for file in os.listdir("Kardexs")]
    listDataWord=[]
    for file in files:
        dWord=dataWord(file)
        # llamar a mi funcion de busqueda de nombres naturales o juridicos
        listDataWord.append(dWord.getdataword())
    print(listDataWord)

sub_main()