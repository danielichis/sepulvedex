from muttlib.gsheetsconn import GSheetsClient
import pandas as pd
from utilities import pathsManager
import numpy as np
from pathlib import Path
from datetime import datetime
import os
pm=pathsManager().currentFolderPath

credsPath=os.path.join(pm,"src","bot sepulveda.json")
GOOGLE_SHEETS_SECRETS_JSON_FP = Path(credsPath)
# id de del google sheet
GSHEETS_SPREAD_ID = '1jeGPabSE47eSLAGAR7OdKrzy7Dg1gaf222iVPwzi-Bc'
#nombre de la hoja 
GSHEETS_WORKSHEET_NAME = 'Hoja 1'
#nombre hoja2
sheet2_name='Correos'


def findEmail(confrontadora,listaCorreos):
    for correo in listaCorreos:
        if confrontadora==correo[0]:
            return correo[1]
    return "no encontrado"

def get_data_from_gsheet():
    gsheets_client = GSheetsClient(GOOGLE_SHEETS_SECRETS_JSON_FP)
    spreadSheet_id, worksheet = (GSHEETS_SPREAD_ID, GSHEETS_WORKSHEET_NAME)
    #gsheets_client.insert_from_frame(df, spreadSheet_id, index = False, worksheet = worksheet, first_cell_loc= 'C3', header = False,preclean_sheet=False)
    return_df = gsheets_client.to_frame(spreadSheet_id, worksheet = worksheet)
    #print("primera consulta hecha")
    return_df2 = gsheets_client.to_frame(spreadSheet_id, worksheet = sheet2_name)
    #print("segunda consulta hecha")
    #print(return_df)
    listkardexs=return_df[['Kardex','Confrontado (SI/ NO/ EN PROCESO)',"Confrontador/a"]].values.tolist()
    listaCorreos=return_df2.values.tolist()
    listToDownload=[]
    for i,kardex in enumerate(listkardexs):   
        if kardex[1] == 'EN PROCESO (BOT)':
            dictkardex={
                "rowNumber":i+2,
                "kardex":kardex[0],
                "confrontado":kardex[1],
                "confrotador":kardex[2],
                "correo":findEmail(kardex[2],listaCorreos)
                }
            listToDownload.append(dictkardex)
    n=len(listToDownload)
    listToupdate=['BOT EN EJECUCION ...']*n
    df=pd.DataFrame(listToupdate)
    #get the current date and time
    now = str(datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
    cell='E'+str(listToDownload[0]['rowNumber'])
    cell2='O'+str(listToDownload[0]['rowNumber'])
    listToupdate2=[now]*n
    df2=pd.DataFrame(listToupdate2)
    gsheets_client.insert_from_frame(df, spreadSheet_id, index = False, worksheet = worksheet, first_cell_loc= cell, header = False,preclean_sheet=False)
    gsheets_client.insert_from_frame(df2, spreadSheet_id, index = False, worksheet = worksheet, first_cell_loc= cell2, header = False,preclean_sheet=False)
    #print(listToDownload)
    return listToDownload

def updateSheet(listToDownload):
    gsheets_client = GSheetsClient(GOOGLE_SHEETS_SECRETS_JSON_FP)
    spreadSheet_id, worksheet = (GSHEETS_SPREAD_ID, GSHEETS_WORKSHEET_NAME)
    #gsheets_client.insert_from_frame(df, spreadSheet_id, index = False, worksheet = worksheet, first_cell_loc= 'C3', header = False,preclean_sheet=False)
    return_df = gsheets_client.to_frame(spreadSheet_id, worksheet = worksheet)
    #print("primera consulta hecha")
    n=len(listToDownload)
    listToupdate=['SI']*n
    df=pd.DataFrame(listToupdate)
    #get the current date and time
    now = str(datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
    cell='E'+str(listToDownload[0]['rowNumber'])
    cell2='O'+str(listToDownload[0]['rowNumber'])
    listToupdate2=[now]*n
    df2=pd.DataFrame(listToupdate2)
    gsheets_client.insert_from_frame(df, spreadSheet_id, index = False, worksheet = worksheet, first_cell_loc= cell, header = False,preclean_sheet=False)
    gsheets_client.insert_from_frame(df2, spreadSheet_id, index = False, worksheet = worksheet, first_cell_loc= cell2, header = False,preclean_sheet=False)
    
#updateSheet()
    #gsheets_client.insert_from_frame(df, spreadSheet_id, index = False, worksheet = worksheet, first_cell_loc= 'C3', header = False,preclean_sheet=False)
#print(get_data_from_gsheet())
# while True:
#     print('getting data from google sheet...')
#     get_data_from_gsheet()
#     time.sleep(20)
