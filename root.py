from re import M
from time import time
import requests
import json
import openpyxl
import string
import datetime
import os

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from datetime import date, datetime, timedelta

# GOOGLE AUTH 

import os.path
from googleapiclient.discovery import build
from google.oauth2 import service_account

SERVICE_ACCOUNT_FILE = 'keys.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

creds = None
creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# The ID of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '177LVKVaM5a-7tokmpp8anueSLgyBN2DsYvlVA-1CKM8'



service = build('sheets', 'v4', credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()



#In 2D array speichern oder gibt es andere mlöglichkeit?






#print(request)

#Ende

wb = load_workbook('Daten.xlsx') 

ws = wb.active

dataName_array = ["Currencies", "Energies" , "Equities" , "Financials" , "Grains" , "Meats" , "Metals" , "Softs"]

b = 0

progress = 0.0

isCme = False

datum = (datetime.now() - timedelta(1))
datum_weekday = datum.weekday()

if(datum_weekday < 5): 
    tradeDate = datum.strftime('%Y%m%d')
elif(datum_weekday == 5):
    tradeDate = (datum - timedelta(1)).strftime('%Y%m%d')
elif(datum_weekday == 6):
    tradeDate = (datum - timedelta(2)).strftime('%Y%m%d')

#Get trade Date from Website




headers = {
    "Accept": "application/json",
    "Accept-Encoding": "gzip, deflate",
    "User-Agent": "Mozilla/5.0"
}

cord_col_a = 1
cord_col_b = 2 


def printProgressBar(progress):

    print("[", end="")

    for i in range(0 , 100):
        if(i < progress): print("=", end="")
        if(i > progress): print("-" , end="")

    
    print("]")
    


print(            """\n                                                                     
                                    .......                                               
                              :::::=@@@@@@@-::::.                                         
                          .::-%@@%%%*******#%@@@#::::::::::::::::::.                      
                      .---*%%%******************#%%%%%%%%####%%%%%%*-------               
                    .=*%%%#******#%%%#*****************************#%%%%%%%===-           
                   =+%#*********%#*#@%%#***********************************%%%#=-         
                 ++%#***********@@+*#%@%****************####***##########*****#%#+-       
                 @@*************%%@@%%%#************#####%%%########%%%%@#******#%#+:     
               **@@*************##%%%%#*************#####%%%%%%###*###%%%%%*******%@=     
             +#****%%%%@@@%#******###***********#####***##%%%##%%%###%%%%@@***###*####:   
           +#*+====+++++++#%@@##*********+++*****##%######%###%%%%%%%%%%%%###*#%#*##@@-   
         =#*+====*****#%%%@@@@@@####****+***###****##%%%%###%%%%%%%%%%%%%%*%#+#%#*#%##%%. 
       =%*+==+*****@@@#+++++=--=@@@@%##%###%%%%####***####%%%%%%%%%%%%%%%%##%@%*###*#%@@: 
     -@#===+*****@%-=+-.:---%@@@**@@%###*****%#%%%%%%%#%%%%%%%%#%%%%%%%##%%%%@@*#%#*#%@@: 
     -@#=+***%@@@--:-+=-*@@@+=**@@%%%%%%%#░█████╗░██████╗░░█████╗░░██╗░░░░░░░██╗░░░*#%#*@@
   :@%=+***%@+===++**@@@*==+****@@%%%%%%%%██╔══██╗██╔══██╗██╔══██╗░██║░░██╗░░██║░░░%%%##@@
  :-%#++*#%*+::--=+%#++++*****@@%%%%%%%%%%██║░░╚═╝██████╔╝███████║░╚██╗████╗██╔╝░░░%%%%%@@
 .@%=+##%+:=+=+%%%#++****#########%%%%#%%%██║░░██╗██╔══██╗██╔══██║░░████╔═████║░░░░%%%%%@@
--%#**%*=-:+*#%++++**###****.    .***#%%%%╚█████╔╝██║░░██║██║░░██║░░╚██╔╝░╚██╔╝░██╗@%%%%@@
@@=+@#===+###++****##**+:            =@%##░╚════╝░╚═╝░░╚═╝╚═╝░░╚═╝░░░╚═╝░░░╚═╝░░╚═╝@%%%%@@
@@###+:=**#*+**####+-                :+*##%%%%%%%%%%%%%%%####%%%%%%%%%%%%###%%%%%%%@%%%%@@
===-:-**#*+#%%%===-                    +@%#%%%#%%%%%%%#%%%%%%%%#%%%%%%%###%%%%%%%@@%@@@@+=
   .***##%%+---                        :-#%##%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%@@@@%%%@@+-  
   :@@%+---.                             :-#%%%%%%%%%%%%%%%#%%%%%%%%%%%%%%%@@@@%%%%%@@-   
    :::.                                   ::@@@@@%%%%%%%%%%%%%%%%%%%%%@@@@%%@@%%%%%@@-   
                                             @@##%%@@@@@@@@@@@@@@@@@@@@%%%%%%@@%%%%%@@-   
                                             @@##%%@*.*@%%%%%#%%%%%%%@@%%%%@@%%%%%%%%%%%. 
                                           *%@@%%@*     #@#*##%%%%%%%%%@@@%%##%%%%%%%%@@: 
                                           *@**@#         #@*#%%%%#%%%%@@*=%#*%%%@@@%%@@: 
                                         #@+=@@           #@###%%###%%%@@**@@%%%%@+ +@@@. 
                                        .*@#*%#             @@*#@@%####@@+=@@%%%%@+       
                                       +@*+%@               ####@* -###@@+=@@##%%@+       
                                     .:*%####                 ###+   =#@@+=@% .###=       
                                     =@#+%@.                           #@++@@             
                                   .-+###**.                           #@**@@             
                                  =+#*+#@-                             #@**@%             
                              =+++#*++*%@-                         :+++*#**@%             
                            =+**###*+*%##*+.                     .+*#@@+=####+=           
                            %@==%%@%+#@#*%@.                   .***#%@@*+@@**@#           
                            %@%#-=@@%*-*%@@.                   :@@%*-#@@%--%%@#           
                            ::::  :::. .:::                     :::. .:::  :::.           
                                """ )

inc = 4
ch = 'B'

for b in range (0, 8):

    inc = 4
    

    if b >= 1: 
        cord_col_a = cord_col_a + 3
        cord_col_b = cord_col_b + 3

    with open("data/" + dataName_array[b] + ".json") as f:
        info_data = json.load(f)

    ws[get_column_letter(cord_col_a) + "2" ] = dataName_array[b]

    cord_a = 4
    cord_b = 5

    print("Fetching: " + dataName_array[b] + ".json")

    #Momentan preloaded: Man könnte auch je eintrag in for l schleife prozentzahl addieren!

    printProgressBar(progress)

    print("Progress: " + str(progress) + "%\n")

    

    for l in range(0 , len(info_data["infoData"])):

        rangeS = "Test!" + str(ch) + str(inc) #Name der Tabelle! 
        
        
        progress = round(progress + (12.5/len(info_data["infoData"])), 2)


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

        if(info_data["infoData"][l]["from"] == "cme"): #Ist cme url?

            params = {
                "tradeDate": tradeDate, #wie verändert sich das trade datum?
                "pageSize": "50",
                "_": "1620683546888"
            }

            url_id = (info_data["infoData"][l]["url-id"])

            url = "https://www.cmegroup.com/CmeWS/mvc/Volume/Details/F/"+ url_id +"/"+ tradeDate + "/P"
            print(url)
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
        name = info_data["infoData"][l]["name"]
        i = 0 #für schleife monate reset

        aoa = [[name], ["MONAT", "TOTAL"]]
       
        

        cord_b = cord_b + 1

        while i < 3:

            if(isCme == True):
                if i == len(data_Cme["monthData"]): #Exit wenn nicht mehr monate # Dann im array leer!
                    cord_b = cord_b + 1
                    break
 
                ws[get_column_letter(cord_col_a) + str(cord_b)] = data_Cme["monthData"][i]["month"]
                ws[get_column_letter(cord_col_b) + str(cord_b)] = data_Cme["monthData"][i]["totalVolume"]

                month = data_Cme["monthData"][i]["month"]
                totalVolume = data_Cme["monthData"][i]["totalVolume"]
                arrayMonat = [month, totalVolume]
                aoa.append(arrayMonat)

                request = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, 
                                range=rangeS ,valueInputOption="USER_ENTERED", body={"values": aoa}).execute()
        

            if(isCme == False):

                if i == len(data_ice): #Bug fix "Out of Range z.190"
                    cord_b = cord_b + 1
                    break
                
                ws[get_column_letter(cord_col_a) + str(cord_b)] = data_ice[i]["marketStrip"]
                ws[get_column_letter(cord_col_b) + str(cord_b)] = data_ice[i]["volume"]

                month = data_ice[i]["marketStrip"]
                totalVolume = data_ice[i]["volume"]
                arrayMonat = [month, totalVolume]
                aoa.append(arrayMonat)

                request = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, 
                                range=rangeS ,valueInputOption="USER_ENTERED", body={"values": aoa}).execute()

            i = i + 1 #Increment
           


            cord_b = cord_b + 1

        inc = inc + 6

        
        #Abstand zwischen neuen Datensätzen:

        cord_b = cord_b + 2
        cord_a = cord_a + 6

        
    
        i = 0 #reset für nächsten durchlauf

    ch = chr(ord(ch) + 3)
    wb.save("Daten.xlsx")



print(""" \n 
███████╗██╗███╗░░██╗██╗░██████╗██╗░░██╗██╗
██╔════╝██║████╗░██║██║██╔════╝██║░░██║██║
█████╗░░██║██╔██╗██║██║╚█████╗░███████║██║
██╔══╝░░██║██║╚████║██║░╚═══██╗██╔══██║╚═╝
██║░░░░░██║██║░╚███║██║██████╔╝██║░░██║██╗
╚═╝░░░░░╚═╝╚═╝░░╚══╝╚═╝╚═════╝░╚═╝░░╚═╝╚═╝
""")

os.system('pause')