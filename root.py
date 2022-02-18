import requests
import json
import openpyxl
import string

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

wb = load_workbook('Daten.xlsx') 
ws = wb.active

ws.delete_rows(1, ws.max_row+1) ## ACHTUNG: LÖSCHT ALLE DATEN DIE BEREITS IN TABELLE STEHEN!

with open("info.json") as f:
    info_data = json.load(f)

with open("infotest.json") as b: 
    infotest_data = json.load(b)

b = 0

params = {
    "tradeDate": "20211210", #wie verändert sich das trade datum?
    "pageSize": "50",
    "_": "1620683546888"
}

headers = {
    "Accept": "application/json",
    "Accept-Encoding": "gzip, deflate",
    "User-Agent": "Mozilla/5.0"
}

cord_a = 2;
cord_b = 4;
cord_c = 5;

print ()


for b in range (0,len(infotest_data["equeties"])):
    url_id = (info_data["infoData"][b]["url-id"])
    #Test

    url = "https://www.cmegroup.com/CmeWS/mvc/Volume/Details/F/"+ url_id +"/20211210/P"

    response = requests.get(url, params=params, headers=headers) #URL
    response.raise_for_status()

    data = response.json() #Data von Url Json

    ws["A" + str(cord_a)].fill = PatternFill(start_color="fd3535", end_color="cc0f0f", fill_type = "solid")
    ws["A" + str(cord_a)] = info_data["infoData"][b]["name"]

    i = 0 #für schleife monate reset

    while i < 3: #Monate Schleife

        ws["A" + str(cord_b)] = "Monat:"
        ws["B" + str(cord_b)] = data["monthData"][i]["month"]

        ws["A" + str(cord_c)] = "Total:"
        ws["B" + str(cord_c)] = data["monthData"][i]["totalVolume"]

        cord_b = cord_b + 3 #Abstände zwischen den monaten
        cord_c = cord_c + 3 #Abstände zwischen Total

        i = i + 1 

    ##Abstand zwischen neuen Datensätzen

    cord_b = cord_b + 2
    cord_a = cord_c - 1 ##cord a beim 2 durchlauf soll 2 höher sein wie cord c abstand zu nächstem eintrag
    cord_c = cord_c + 2
    

    i = 0 #reset für nächsten durchlauf

    wb.save("Daten.xlsx")
    