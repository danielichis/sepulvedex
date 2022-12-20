from getwords import download_legacy
from scrapDocxs import GettingkardexData
from scrapPages import scrapingPages

def main():
    download_legacy()
    GettingkardexData()
    scrpy=scrapingPages()
    scrpy.gettingPagesData()
main()
