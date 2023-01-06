from getwords import download_legacy
from scrapDocxs import GettingkardexData
from scrapPages import scrapingPages
from validate import readJsonPages
from validateMore import validateCorrelatives
from amountsValidation import amountsValidation0
from sendEmails import sendEmails
def main():
    download_legacy()
    GettingkardexData()
    scrpy=scrapingPages()
    scrpy.gettingPagesData()
    readJsonPages()
    validateCorrelatives()
    amountsValidation0()
    x = sendEmails()
    x.send()
    print("-----------------------------------------PROCESO TERMINADO ............................")
main()

