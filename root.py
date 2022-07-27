from cgitb import text
from email import message
from email.policy import HTTP
from re import M
from time import time
from tkinter import W
from urllib.error import HTTPError
from xml.dom.minidom import Document
import requests
import json
import datetime
import sched, time
import logging


import telegram
import constants as keys
import responses as R

from datetime import date, datetime, timedelta
#Google Auth for Sheets API:
from googleapiclient.discovery import build
from google.oauth2 import service_account
from telegram.ext import * 

##Dependencies##
#telegram-bot
#requests
#openpyxl #Für excel
#googleapiclient  pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib


def send_file(file, chat_id):
    bot = telegram.Bot(keys.API_KEY)

    with open(file, 'r') as f:
        bot.send_document(chat_id = chat_id, document = f )
        f.close()


def delete_logfile(file):
    with open(file, 'a') as f:
        f.truncate(0)



# Telegram Bot Functions #

def start_command(update, context):
    update.message.reply_text("Willkommen beim Craw-Bot. Für Hilfe: /help")

def help_command(update, context):
    update.message.reply_text("/log : Gibt die Log-Datei zurück. \n/craw : Startet das Script \n/clear : Löscht die log-Datei ")

def log_command(update, context):
    chat_id = update.effective_chat.id
    send_file('logCraw.log',chat_id)

def clear_command(update, context):
    update.message.reply_text('Log Datei gelöscht!')
    delete_logfile('logCraw.log')
    

def hanlde_message(update, context):
    text = str(update.message.text).lower()
    response = R.sample_responses(text)
    update.message.reply_text(response)

def craw_command(update, context):
    update.message.reply_text("Gestartet")
    main(s, True)

def error(update, context):
    print(f"Update {update} caused error {context.error}")

######################

# Call the Sheets API #

SERVICE_ACCOUNT_FILE = 'keys.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

creds = None
creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# Online # 
#SAMPLE_SPREADSHEET_ID = keys.SAMPLE_SPREADSHEET_ID
##########

# TEST #
SAMPLE_SPREADSHEET_ID = keys.SAMPLE_SPREADSHEET_ID_TEST
########
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
    # Datum is Weekday
    tradeDate = datum.strftime('%Y%m%d')
elif(datum_weekday == 5):
    # If Datum = Saturday -> Datum - OneDay
    tradeDate = (datum - timedelta(1)).strftime('%Y%m%d')
elif(datum_weekday == 6):
    # IF Datum = Sunday -> Datum - 2 Days
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

    ###########
    # This function lowers the trade date if there is no data 
    # Attention: param tradeDate is already Weekday
    # So the function checks when the last update has been and sets the tradeDate;
    ###########

    #While Array Len == 0 -> tradeDate nicht richtig -> Minus as long as tradeDate_data is empty
    #While Array Len > 0 -> tradeDate richtig

    while(len(tradeDate_data["monthData"]) == 0):
        tradeDateInt = int(tradeDate) - 1
        tradeDate = str(tradeDateInt)
        
        newUrl = "https://www.cmegroup.com/CmeWS/mvc/Volume/Details/F/"+ url_id +"/"+ tradeDate + "/P"
        try: #CME-Group
            response = requests.get(newUrl, params=params, headers=headers) #URL
            response.raise_for_status()
        except requests.exceptions.Timeout:
            logging.exception("Exception occured (timeout) - Connect to CME - CheckTradeDate")
            print("time-out")
        except requests.exceptions.ConnectionError:
            logging.exception("Exception occured (conn err) - Connect to CME - CheckTradeDate")
            print('Connection Error')
        except requests.exceptions.HTTPError:
            logging.exception("Exception occured (Bad Gateway) - Test? -CheckTradeDate")
            print("BadGateway - checkTradeDate()")
        tradeDate_data = response.json() #Data von Url Json
   
    return tradeDate_data
    
def printProgressBar(progress):

    ####
    # Function that prints the Progress Bar
    ####

    print("[", end="")

    for i in range(0 , 100):
        if(i < progress): print("=", end="")
        if(i > progress): print("-" , end="")

    
    print("]")
    
def botrun():
    ### Bot Function ###
    updater = Updater(keys.API_KEY, use_context=True)
    dp = updater.dispatcher

    dp.add_handler(CommandHandler("start", start_command))
    dp.add_handler(CommandHandler("help", help_command))
    dp.add_handler(CommandHandler("craw", craw_command))
    dp.add_handler(CommandHandler("log", log_command ))
    dp.add_handler(CommandHandler("clear", clear_command ))

    dp.add_handler(MessageHandler(Filters.text, hanlde_message))

    dp.add_error_handler(error)

    updater.start_polling(5)
    ######

def main(sc , param = False):
    #bot = telegram.Bot(keys.API_KEY)

    ############
    # Main Function
    # Traversing all Datasets (/data/...json)
    ############


    logging.info("Start Process CRAW")
    #bot.send_message(chat_id="2143240853" ,text="Craw gestartet!")
    dataName_array = ["Currencies", "Energies" , "Equities" , "Financials" , "Grains" , "Meats" , "Metals" , "Softs"]

    #### Standard Werte###
    b = 0
    progress = 0.0
    isCme = False
    ######################

    BeautifulPrintouts("start")
    
    inc = 4 #Increment in Google Sheets??
    ch = 'B' #Used for GoogleSheets Table

    #For Schleife -> Gehe durch Datensätze (in Data ordner)
    for b in range (0, 8):

        inc = 4 #Excel

        with open("data/" + dataName_array[b] + ".json") as f:
            info_data = json.load(f)

        print("Fetching: " + dataName_array[b] + ".json")

        printProgressBar(progress)

        print("Progress: " + str(progress) + "%\n")


        for l in range(0 , len(info_data["infoData"])):
            #####
            # Traversing through infoData from .json Datasets
            #####

            rangeS = "Kontraktvolumen!" + str(ch) + str(inc) #Name der Tabelle auf GoogleSheets! 
            
            progress = round(progress + (12.5/len(info_data["infoData"])), 2)

            if(info_data["infoData"][l]["from"] == "theice"): 
                #####################
                # Ist TheIce url?
                # GetData safe in data_ice variable
                #####################

                isCme = False

                params_ice = {
                    "getContractsAsJson": "",
                    "productId": info_data["infoData"][l]["url-id"], 
                    "hubId": info_data["infoData"][l]["hub-id"], 
                }   

                url_ice = "https://www.theice.com/marketdata/DelayedMarkets.shtml?"

                try: #TheIce
                    response_Ice = requests.get(url_ice, params=params_ice, headers=headers)
                    response_Ice.raise_for_status()
                except requests.exceptions.Timeout:
                    logging.exception("Exception occured - Timeout - ICE - Main()")
                    print("timeOut - TheICE - Main()")
                except requests.exceptions.ConnectionError:
                    logging.exception("Exception occured (conn err) - Connect to TheICE - Main()")
                    print("ConnErr - TheIce - Main()")
                except requests.exceptions.HTTPError:
                    logging.exception("Exception occured (Bad Gateway) - Test? - TheICE - Main()")
                    print("BadGateway - TheICE - Main()")

                data_ice = response_Ice.json() #Data von Url Json

            if(info_data["infoData"][l]["from"] == "cme"): 
                #####################
                # Ist cme url?
                # Get Data and safe into DataCME variable
                # Calls CheckTradeDate Function to prevent saving empty Dataset to variable
                #####################

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
                    response.raise_for_status()
                except requests.exceptions.Timeout:
                    logging.exception("Exception occured (timeout) - Connect to CME - Main()")
                    print("time-out")
                except requests.exceptions.ConnectionError:
                    logging.exception("Exception occured (conn err) - Connect to CME - Main()")
                    print('Connection Error')
                except requests.exceptions.HTTPError:
                    logging.exception("Exception occured (Bad Gateway) - Test? - CME - Main()")
                    print("BadGateway - CME - Main()")
                
                data_Cme_NC = response.json() #Data von Url Json

                data_Cme = checkTradeDate(data_Cme_NC, tradeDate, url_id, params)

                isCme = True

            name = info_data["infoData"][l]["name"]
            i = 0 #für schleife monate reset
            aoa = [[name], ["MONAT", "TOTAL"]] # Array for Data -> initialized with Name, l.377,387 append Data of ICE/CME

            while i < 5:
                # Traversing through Data of ICE/CME
                if(isCme == True):
                    if i == len(data_Cme["monthData"]): #Exit wenn nicht mehr monate im Array
                        break

                    month = data_Cme["monthData"][i]["month"]
                    totalVolume = data_Cme["monthData"][i]["totalVolume"]
                    arrayMonat = [month, totalVolume]
                    aoa.append(arrayMonat) # Apend Month to aoa Array (Array of Data)

                    
                if(isCme == False):
                    if i == len(data_ice): #Exit wenn nicht mehr monate im Array 
                        break

                    month = data_ice[i]["marketStrip"]
                    totalVolume = data_ice[i]["volume"]
                    arrayMonat = [month, totalVolume]
                    aoa.append(arrayMonat) # Append Month to aoa Array (Array of Data)


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
    #bot.send_message(chat_id="2143240853" ,text="Craw beendet!")
    logging.info("Finished Process CRAW")
    if(param == False):
        s.enter(3600, 1, main, (sc,)) #Run every hour
    
    

botrun()
##Start after 5 SEC ###############
s.enter(5, 1, main, (s,False))
s.run()
#################################

#Notifiyer wenn Crash (only in log rn)
#BOT
    #Clear Log
    #Last run
    #Next run?
