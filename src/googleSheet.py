from muttlib.gsheetsconn import GSheetsClient
import pandas as pd
from utilities import pathsManager
import numpy as np
from pathlib import Path
import time
import os
pm=pathsManager().currentFolderPath

credsPath=os.path.join(pm,"src","bot sepulveda.json")
GOOGLE_SHEETS_SECRETS_JSON_FP = Path(credsPath)
# id de del google sheet
GSHEETS_SPREAD_ID = '1jeGPabSE47eSLAGAR7OdKrzy7Dg1gaf222iVPwzi-Bc'
#nombre de la hoja 
GSHEETS_WORKSHEET_NAME = 'Hoja 1'



def get_data_from_gsheet():
    gsheets_client = GSheetsClient(GOOGLE_SHEETS_SECRETS_JSON_FP)
    spreadSheet_id, worksheet = (GSHEETS_SPREAD_ID, GSHEETS_WORKSHEET_NAME)

    #gsheets_client.insert_from_frame(df, spreadSheet_id, index = False, worksheet = worksheet, first_cell_loc= 'C3', header = False,preclean_sheet=False)

    return_df = gsheets_client.to_frame(spreadSheet_id, worksheet = worksheet)
    #print(return_df)
    listkardexs=return_df[['Kardex','Confrontado (SI/ NO/ EN PROCESO)']].values.tolist()
    listToDownload=[]
    for kardex in listkardexs:   
        if kardex[1] == 'NO':
            listToDownload.append(kardex[0])
            print(kardex[0])
    return listToDownload

#print(get_data_from_gsheet())
# while True:
#     print('getting data from google sheet...')
#     get_data_from_gsheet()
#     time.sleep(20)
