import requests
url="https://httpbin.org/headers"

headers={
    "User-agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36",
    "Acept-Language":"es-GB,en;q=0.5",
    "Referer":"https://google.com",
    "DNT":"1",
    "Daniel":"please dont block me"
}
r=requests.get(url,headers=headers)
print(r.text)

