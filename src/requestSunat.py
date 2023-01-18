import requests
from bs4 import BeautifulSoup
Urlsunat="https://www.amazon.com/-/es/Bater%C3%ADas-alcalinas-rendimiento-AAA-paquete/dp/B00LH3DMUO/ref=lp_16225009011_1_2?th=1"

r=requests.get(Urlsunat)
print(r.text)
