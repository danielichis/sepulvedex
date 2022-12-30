import docx
import re
from numerToLetters import numero_a_moneda_sunat
from validate import style_Token,split_Runs
def validate_more(alltext,doc):
    # Getting more info
    monts=re.findall(r"([0-9,.]*) \(([\w\s]* [Y-y]? ?00/100)",alltext)
    for mont in monts:
        y=mont[1].replace("CON","Y")
        numero=mont[0].replace(",","")
        x=numero_a_moneda_sunat(float(mont[0].replace(",",""))).replace(" SOLES","")
        x=x.replace("CON","Y")
        if x!=y:
            split_Runs(doc, numero)
            style_Token(doc, numero,True)
