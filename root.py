import requests
import json
import openpyxl
import string
import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta

wb = load_workbook('Daten.xlsx') 
ws = wb.active

dataName_array = ["Currencies", "Energies" , "Equities" , "Financials" , "Grains" , "Meats" , "Metals" , "Softs"]

# ws.delete_rows(1, ws.max_row+1) ## ACHTUNG: LÖSCHT ALLE DATEN DIE BEREITS IN TABELLE STEHEN!

b = 0

progress = 0

isCme = False

tradeDate = (datetime.now() - timedelta(1)).strftime('%Y%m%d')

print(tradeDate)

params = {
    "tradeDate": "20211216", #wie verändert sich das trade datum?
    "pageSize": "50",
    "_": "1620683546888"
}

headers = {
    "Accept": "application/json",
    "Accept-Encoding": "gzip, deflate",
    "User-Agent": "Mozilla/5.0"
}

cord_col_a = 1
cord_col_b = 2 


for b in range (0, 8):

    if b >= 1: 
        cord_col_a = cord_col_a + 3
        cord_col_b = cord_col_b + 3

    with open("data/" + dataName_array[b] + ".json") as f:
        info_data = json.load(f)

    ws[get_column_letter(cord_col_a) + "2" ] = dataName_array[b]

    cord_a = 4
    cord_b = 5

    print("data/" + dataName_array[b] + ".json")

    #Momentan preloaded: Man könnte auch je eintrag in for l schleife prozentzahl addieren!

    #progress = progress + 12,5

    print("Progress: " + str(progress) + "%")

    for l in range(0 , len(info_data["infoData"])):

        if(info_data["infoData"][l]["from"] == "theice"): #Ist TheIce url?
            
            isCme = False

            params_ice = {
                "getContractsAsJson": "",
                "productId": info_data["infoData"][l]["url-id"], 
                "hubId": info_data["infoData"][l]["hub-id"], 
            }   

            url_ice = "https://www.theice.com/marketdata/DelayedMarkets.shtml?"

            try: #TheIce
                response_Ice = requests.get(url_ice, params=params_ice, headers=headers)
            except requests.exceptions.RequestException as e1:
                raise SystemExit(e1)

            response_Ice.raise_for_status()
            data_ice = response_Ice.json() #Data von Url Json
            print(response_Ice.url)

        if(info_data["infoData"][l]["from"] == "cme"): #Ist cme url?

            url_id = (info_data["infoData"][l]["url-id"])

            url = "https://www.cmegroup.com/CmeWS/mvc/Volume/Details/F/"+ url_id +"/"+ tradeDate + "/F"

            try: #CME-Group
                response = requests.get(url, params=params, headers=headers) #URL
            except requests.exceptions.RequestException as e:
                raise SystemExit(e)

            response.raise_for_status()

            data_Cme = response.json() #Data von Url Json

            isCme = True
         

        ws[get_column_letter(cord_col_a) + str(cord_a)] = info_data["infoData"][l]["name"]
        ws[get_column_letter(cord_col_a) + str(cord_b)] = "Monat:"
        ws[get_column_letter(cord_col_b) + str(cord_b)] = "Total Volume:"

        i = 0 #für schleife monate reset

        cord_b = cord_b + 1

        while i < 3: #Monate Schleife

            if(isCme == True):
                ws[get_column_letter(cord_col_a) + str(cord_b)] = data_Cme["monthData"][i]["month"]

                ws[get_column_letter(cord_col_b) + str(cord_b)] = data_Cme["monthData"][i]["totalVolume"]

            if(isCme == False):
                ws[get_column_letter(cord_col_a) + str(cord_b)] = data_ice[i]["marketStrip"]

                ws[get_column_letter(cord_col_b) + str(cord_b)] = data_ice[i]["volume"]

            i = i + 1 #Increment

            cord_b = cord_b + 1

        #Abstand zwischen neuen Datensätzen:

        cord_b = cord_b + 2
        cord_a = cord_a + 6
    
        i = 0 #reset für nächsten durchlauf

    wb.save("Daten.xlsx")
    