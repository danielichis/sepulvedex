from getwords import download_legacy
from scrapDocxs import GettingkardexData
from scrapPages import scrapingPages
from validate import readJsonPages
from validateMore import validateCorrelatives
def main():
    download_legacy()
    GettingkardexData()
    scrpy=scrapingPages()
    scrpy.gettingPagesData()
    readJsonPages()
    validateCorrelatives()
main()

