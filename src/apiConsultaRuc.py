import requests
import json
import random
def getRuc(ruc):
    url = "https://www.softwarelion.xyz/api/sunat/consulta-ruc"
    tokend="Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1c2VyX2lkIjozNDU3LCJjb3JyZW8iOiJkY2hhY29uYkB1bmkucGUiLCJpYXQiOjE2NzQxMDI2NDF9.s05kgg-A36WQoJfXrtbkEB3qIfSwOy9XbU1mGVUBiEc"
    tokensp="Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1c2VyX2lkIjozNDU4LCJjb3JyZW8iOiJib3RAbm90YXJpYXJvc2FsZXNzZXB1bHZlZGEuY29tLnBlIiwiaWF0IjoxNjc0MTA3MDMxfQ.hne_DSYhU7msnGCnlSzhStGOd4G0QAs9TlLllita-SY"
    
    
    listokens=[tokend,tokensp]
    #random token from listtokens
    token=random.choice(listokens)
    _json = {"ruc": ruc}
    _headers = {"Content-type": "application/json", "Authorization": token}

    response=requests.post(url, data=json.dumps(_json), headers=_headers)

    dataJson = response.json()
    if dataJson["success"]==False:
        return "Ruc Not Found"
    else:
        emrpresa=dataJson["result"]
        return emrpresa["razonSocial"]

# r=getRuc("20256149681")
# print(r)