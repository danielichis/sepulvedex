from getwords import download_legacy
from scrapDocxs import GettingkardexData
from scrapPages import scrapingPages
from numerToLetters import numero_a_moneda_sunat
import re
def main():
    download_legacy()
    GettingkardexData()
    scrpy=scrapingPages()
    scrpy.gettingPagesData()
main()

