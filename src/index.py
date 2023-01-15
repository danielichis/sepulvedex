from getwords import download_legacy
from scrapDocxs import GettingkardexData
from scrapPages import scrapingPages
from validate import readJsonPages
from validateMore import validateCorrelatives
from amountsValidation import amountsValidation0
from sendEmails import sendEmails
def main():
    ks=download_legacy() #1 descargar words de legacy a la carpeta kardexs
    GettingkardexData()#2 obtener los datos de los words y guardarlos en dataofWords.json
    scrpy=scrapingPages()#3 obtener los datos de las paginas y guardarlos en dataofPages.json
    scrpy.gettingPagesData()#3 obtener los datos de las paginas y guardarlos en dataofPages.json
    readJsonPages() #4 validar los datos de dataofPages.json
    validateCorrelatives() #5 validar correlativos
    amountsValidation0() #6 validar montos
    x = sendEmails() #7 enviar correos
    x.sendEmails(ks)#7 enviar correos
    print("-----------------------------------------PROCESO TERMINADO ............................")
main()

