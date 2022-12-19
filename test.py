import json
import os

with open('dataofWords.json') as f:
    #kardexs = list(json.load(f).keys())
    listDataWord = json.load(f)
ks = list(listDataWord.keys())
for i,kardexData in enumerate(list(listDataWord.values())):
    print(f"-------reading karex: {ks[i]}")
    print(kardexData['DNI'])
    print(kardexData['RUC'])
    print(kardexData['partidE'])