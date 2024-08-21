import mysql.connector
from mysql.connector import errorcode
import re
import docx 
import os
import os.path
from docx import Document
import gspread

credentials = {
     "installed":{
    "client_id":"427152031284-lstf22oo911ovohv4il6ldh9k6rhkbum.apps.googleusercontent.com",
     "project_id":"quiz-paracadutismo",
     "auth_uri":"https://accounts.google.com/o/oauth2/auth",
     "token_uri":"https://oauth2.googleapis.com/token",
     "auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs",
     "client_secret":"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",
     "redirect_uris":["http://localhost"]
     
     }
  }

authorized_user = {
  "refresh_token": "1//09c3saVBQDNrbCgYIARAAGAkSNwF-L9Ir9K2Qkg9vECeAjqYOPV4JqX5lLQ9vh7-ICiCSql4XxZvZ9GDSNliGCVRhtkH-H-sJjfw",
  "token_uri": "https://oauth2.googleapis.com/token",
  "client_id": "427152031284-lstf22oo911ovohv4il6ldh9k6rhkbum.apps.googleusercontent.com",
  "client_secret": "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", 
  "scopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
    ], 
  "universe_domain": "googleapis.com", 
  "account": "", 
  "expiry": "2024-06-21T16:01:03.051872Z"
  }

gc, authorized_user = gspread.oauth_from_dict(credentials, authorized_user)

worksheet = gc.open("Prova python")


#Crea un Database MySQL
mydb = mysql.connector.connect(
host="localhost",
user="root",
password="",
database="quiz"
)
print('Database Connesso')

cursor = mydb.cursor()


#Crea un Database MySQL    
def crea_database():
    
    DB_NAME = 'quiz'
    
    mycursor = mydb.cursor()
    try:
        mycursor.execute("CREATE TABLE IF NOT EXISTS domande (id INT AUTO_INCREMENT PRIMARY KEY,value TEXT(65000) NOT NULL)")
        AUTO_INCREMENT = 1
    
    except mysql.connector.Error as err:
        print("Failed creating database: {}".format(err))
        exit(1)
    try:   
        mycursor.execute("CREATE TABLE IF NOT EXISTS risposte_corrette (id INT AUTO_INCREMENT PRIMARY KEY,value TEXT(65000) NOT NULL)")
        AUTO_INCREMENT = 1
    except mysql.connector.Error as err:
        print("Failed creating database: {}".format(err))
        exit(1)
    try:
        mycursor.execute("CREATE TABLE IF NOT EXISTS risposte_errate (id INT AUTO_INCREMENT PRIMARY KEY,value TEXT(65000) NOT NULL)")
        AUTO_INCREMENT = 1
    except mysql.connector.Error as err:
        print("Failed creating database: {}".format(err))
        exit(1)
    
      

# Carica il documento Word e lo stampa riga per riga     
def carica_documento():
    document = Document('quiz.docx')
    testo = ""
    
    for paragrafo in document.paragraphs:
        testo += paragrafo.text + "\n"   
    return testo

#Estrae Domande, Risposte corrette e Risposte Sbagliate dal testo  
def estrai_dati(testo):
    
    domande = []
    risposte_corrette = []
    risposte_errate = []
    
    #Pattern per trovare Domande e Risposte
    pattern_domanda = r'(?<!\b[a-d]\))^[A-ZÀ-ÖØ-Ý].*?\?$'
    pattern_risposta_corretta = r'^[A-ZÀ-ÖØ-Ý0-9].*?\( risposta corretta \)$'
    pattern_risposta_errata = r'^(?![A-ZÀ-ÖØ-Ý0-9].*?\( risposta corretta \)|[A-ZÀ-ÖØ-Ý0-9].*?\?$).*$'
    
    #Trova le domande e le risposte
    domande_match = re.findall(pattern_domanda, testo, re.MULTILINE)
    risposte_corrette_match = re.findall(pattern_risposta_corretta, testo, re.MULTILINE)
    risposte_errate_match = re.findall(pattern_risposta_errata, testo, re.MULTILINE)
    
    #Aggiunge le Domande alla lista
    for domanda in domande_match:
        domande.append(domanda.strip())
        
    #Aggiunge le Risposte Corrette alla lista
    for risposta_corretta in risposte_corrette_match:
        risposte_corrette.append(risposta_corretta.strip())
        
    #Aggiunge le Risposte Sbagliate alla lista
    for risposta_errata in risposte_errate_match:
        risposte_errate.append(risposta_errata.strip())
    
    while ("" in risposte_errate):
        risposte_errate.remove("")
    
    return domande, risposte_corrette, risposte_errate
            
def aggiorna_foglio_google(risposte_corrette,risposte_errate,domande):
    spreadsheet = gc.open("Quiz Paracadutismo")
    worksheet = spreadsheet.sheet1
    
    stringhe_domande = domande
    stringhe_risposte_corrette = risposte_corrette
    stringhe_risposte_errate = risposte_errate
    
    
    for domanda in stringhe_domande:
        worksheet.update_cell(1,domanda)
    for i,risposta_corretta in enumerate(stringhe_risposte_corrette, start=1 ):
        worksheet.update_cell( i+1, 2, risposta_corretta)
    for i,risposta_errata in enumerate(stringhe_risposte_errate, start=0):
        row = (i // 3) + 2  # Calcola la riga, iniziando dalla seconda riga
        col = (i % 3) + 3   # Calcola la colonna (3, 4, 5)
        worksheet.update_cell(row, col, risposta_errata)
 
    return True


def compila_database(risposte_errate):
    
    query = ("INSERT INTO risposte_errate (value) VALUES (%s)")
    
    for risposta_errata in risposte_errate:
        cursor.execute(query, (risposta_errata,))
    
    mydb.commit()
    
    return True

    
crea_database()
testo = carica_documento()  
domande, risposte_corrette, risposte_errate = estrai_dati(testo)
compila_database(risposte_errate)

cursor.close()
mydb.close()

#print(risposte_corrette)
#print(risposte_errate)
#aggiorna_foglio_google(risposte_corrette, risposte_corrette, domande)


# for domanda in domande:                  
    # print(domanda)

#for risposta_errata in risposte_errate:
#    print(risposta_errata)
print(" \n \n Programma eseguito con successo")

