from email import message
from re import M
from time import time
import requests
import json
import datetime
import sched, time
import logging
from datetime import date, datetime, timedelta
#Google Auth for Sheets API:
from googleapiclient.discovery import build
from google.oauth2 import service_account

##Dependencies##

#requests
#openpyxl #Für excel
#googleapiclient  pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib


# Call the Sheets API #

SERVICE_ACCOUNT_FILE = 'keys.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

creds = None
creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)

SAMPLE_SPREADSHEET_ID = '1E5-AezYFSapijKRCwsLUfD16qsq2HTh3f3s-pgGFzNo'
#SAMPLE_SPREADSHEET_ID = '177LVKVaM5a-7tokmpp8anueSLgyBN2DsYvlVA-1CKM8'

service = build('sheets', 'v4', credentials=creds)

sheet = service.spreadsheets()

#######################

#### Logging #####
logging.basicConfig(filename="logCraw.log", level=logging.INFO, format='%(asctime)s - %(message)s')
#################

#### Set trade Date ####
datum = (datetime.now() - timedelta(1))
datum_weekday = datum.weekday()

if(datum_weekday < 5): 
    tradeDate = datum.strftime('%Y%m%d')
elif(datum_weekday == 5):
    tradeDate = (datum - timedelta(1)).strftime('%Y%m%d')
elif(datum_weekday == 6):
    tradeDate = (datum - timedelta(2)).strftime('%Y%m%d')
#########################

headers = {
    "Accept": "application/json",
    "Accept-Encoding": "gzip, deflate",
    "User-Agent": "Mozilla/5.0"
}

logging.info("START craw.py")
s = sched.scheduler(time.time, time.sleep)

def BeautifulPrintouts(forw):

    if(forw == "start"):
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

    if(forw == "end"): 
        print(""" \n 
    ███████╗██╗███╗░░██╗██╗░██████╗██╗░░██╗██╗
    ██╔════╝██║████╗░██║██║██╔════╝██║░░██║██║
    █████╗░░██║██╔██╗██║██║╚█████╗░███████║██║
    ██╔══╝░░██║██║╚████║██║░╚═══██╗██╔══██║╚═╝
    ██║░░░░░██║██║░╚███║██║██████╔╝██║░░██║██╗
    ╚═╝░░░░░╚═╝╚═╝░░╚══╝╚═╝╚═════╝░╚═╝░░╚═╝╚═╝
    """)

def checkTradeDate(tradeDate_data, tradeDate, url_id, params):

    #While Array Len == 0 -> tradeDate nicht richtig
    #Wihile Array Len > 0 -> tradeDate richtig

    while(len(tradeDate_data["monthData"]) == 0):
        tradeDateInt = int(tradeDate) - 1
        tradeDate = str(tradeDateInt)
        
        newUrl = "https://www.cmegroup.com/CmeWS/mvc/Volume/Details/F/"+ url_id +"/"+ tradeDate + "/P"
        try: #CME-Group
            response = requests.get(newUrl, params=params, headers=headers) #URL
        except requests.exceptions.Timeout:
            logging.exception("Exception occured (timeout) - Connect to CME")
            print("time-out")
        except requests.exceptions.ConnectionError:
            logging.exception("Exception occured (conn err) - Connect to CME")
            print('Connection Error')
                

        response.raise_for_status()

        tradeDate_data = response.json() #Data von Url Json

    
        
    return tradeDate_data
    
def printProgressBar(progress):

    print("[", end="")

    for i in range(0 , 100):
        if(i < progress): print("=", end="")
        if(i > progress): print("-" , end="")

    
    print("]")
    
def main(sc):

    logging.info("Start Process CRAW")

    dataName_array = ["Currencies", "Energies" , "Equities" , "Financials" , "Grains" , "Meats" , "Metals" , "Softs"]
    #### Standard Werte###
    b = 0
    progress = 0.0
    isCme = False
    ######################

    BeautifulPrintouts("start")
    
    inc = 4
    ch = 'B' 
    #For Schleife -> Gehe durch Datensätze (in Data ordner)
    for b in range (0, 8):

        inc = 4 

        with open("data/" + dataName_array[b] + ".json") as f:
            info_data = json.load(f)

        print("Fetching: " + dataName_array[b] + ".json")

        printProgressBar(progress)

        print("Progress: " + str(progress) + "%\n")

        for l in range(0 , len(info_data["infoData"])):

            rangeS = "Kontraktvolumen!" + str(ch) + str(inc) #Name der Tabelle auf GoogleSheets! 
            
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
                    logging.exception("Exception occured - Connect to TheICE")
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

                url = "https://www.cmegroup.com/CmeWS/mvc/Volume/Details/F/"+ url_id +"/" + tradeDate + "/P"
                print(url)
                try: #CME-Group
                    response = requests.get(url, params=params, headers=headers) #URL
                except requests.exceptions.Timeout:
                    logging.exception("Exception occured (timeout) - Connect to CME")
                    print("time-out")
                except requests.exceptions.ConnectionError:
                    logging.exception("Exception occured (conn err) - Connect to CME")
                    print('Connection Error')
                

                response.raise_for_status()

                data_Cme_NC = response.json() #Data von Url Json

                data_Cme = checkTradeDate(data_Cme_NC, tradeDate, url_id, params)

                isCme = True

            name = info_data["infoData"][l]["name"]
            i = 0 #für schleife monate reset
            aoa = [[name], ["MONAT", "TOTAL"]]

            while i < 5:

                if(isCme == True):
                    if i == len(data_Cme["monthData"]): #Exit wenn nicht mehr monate # Dann im array leer!
                        break

                    month = data_Cme["monthData"][i]["month"]
                    totalVolume = data_Cme["monthData"][i]["totalVolume"]
                    arrayMonat = [month, totalVolume]
                    aoa.append(arrayMonat)

                    
                if(isCme == False):
                    if i == len(data_ice): #Bug fix "Out of Range z.190"
                        break

                    month = data_ice[i]["marketStrip"]
                    totalVolume = data_ice[i]["volume"]
                    arrayMonat = [month, totalVolume]
                    aoa.append(arrayMonat)


                i = i + 1 #Increment

            inc = inc + 8

            #Write aoa to GoogleSheet
            try:
                request = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, 
                                    range=rangeS ,valueInputOption="USER_ENTERED", body={"values": aoa}).execute()
            except: 
                logging.exception("Connection error - could not be pushed to sheet")
            
            i = 0 #reset für nächsten durchlauf

        ch = chr(ord(ch) + 3) #Increment in GoogleSheet(-> Jeder Datensatz soll nebeneinander stehen B wird zu B+3=E)

    BeautifulPrintouts("end")

    logging.info("Finished Process CRAW")
    s.enter(3600, 1, main, (sc,)) #Run every hour
    
##Start after 5 SEC ###############
s.enter(5, 1, main, (s,))
s.run()
#################################

#Notifiyer wenn Crash
