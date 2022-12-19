from getwords import download_legacy
from scrapDocxs import GettingkardexData
from scrapPages import scrapingPages
from highlightKardexs import highlighting

def main():
    download_legacy()
    GettingkardexData()
    scrpy=scrapingPages()
    scrpy.gettingPagesData()
main()
