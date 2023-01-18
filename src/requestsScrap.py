import requests
from bs4 import BeautifulSoup

sunat_page = 'https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/FrameCriterioBusquedaWeb.jsp'
header = {'user-agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36 Edg/109.0.1518.55'}
response = requests.get(sunat_page, headers=header)
# print(response.status_code)
# print(response.reason)
# print(response.cookies)
html = response.text
soup = BeautifulSoup(html, 'html.parser')
head = soup.find('div', class_ = 'row').find('h1')
print(head)
