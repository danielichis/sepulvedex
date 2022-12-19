import json
  
# Opening JSON file
f = open('dataofPages.json')
  
# returns JSON object as 
# a dictionary
data = json.load(f)
for i in data['emp_details']:
    print(i)
# Closing file
f.close()
# print(data)