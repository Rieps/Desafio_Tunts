from googleapiclient.discovery import build
from google.oauth2 import service_account
import math

#Argument to get the credentials from Keys.json and allow editing of the spreadsheet
SERVICE_ACCOUNT_FILE = 'Keys.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
creds = None
creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('sheets', 'v4', credentials=creds)
sheet = service.spreadsheets()

# The ID of the spreadsheet
SAMPLE_SPREADSHEET_ID = '1JnOVo-uH4Koco__5tQlCivXIqrQ-PPPSpipV_vaohBo'

#Variables for the situations 
AP = [["Aprovado"]]
EF = [["Exame Final"]]
RN = [["Reprovado por nota"]]
  
#Variables to get the value of the grades and absence from the spreadsheet
absence = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,range="engenharia_de_software!C4:C").execute()
falt = absence.get('values', [])

p1 = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,range="engenharia_de_software!D4:D27").execute()
p2 = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,range="engenharia_de_software!E4:E27").execute()
p3 = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,range="engenharia_de_software!F4:F27").execute()

gradep1 = p1.get('values', [])
gradep2 = p2.get('values', [])
gradep3 = p3.get('values', [])
        
#Calculating the average score and printing in the spreadsheet
for i in range(0,len(gradep1)):
    average = math.ceil((int(gradep1[i][0]) + int(gradep2[i][0]) + int(gradep3[i][0])) / 3)
    x=str(i + 4)
    #Comparison of the number of absence to the 25% limit 
    if int(falt[i][0])>15:
                naf = 0
      
                RF = [["Reprovado por falta"]]
                #Command to print in the spreadsheet the RF variable
                request = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, 
                                    range=("engenharia_de_software!G"+x), valueInputOption="USER_ENTERED", body={"values":RF}).execute()
                request = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, 
                                    range=("engenharia_de_software!H"+x), valueInputOption="USER_ENTERED", body={"values":([[str(naf)]])}).execute()
                continue
    #Verification if the Student's average is higher than 70
    if average >= 70:
        naf = 0 
        request = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, 
                            range=("engenharia_de_software!G"+x), valueInputOption="USER_ENTERED", body={"values":AP}).execute()
        request = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, 
                            range=("engenharia_de_software!H"+x), valueInputOption="USER_ENTERED", body={"values":([[str(naf)]])}).execute()
        continue

    #Verification if the Student's average is Below 50
    elif average < 50:
        naf = 0
        request = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, 
                            range=("engenharia_de_software!G"+x), valueInputOption="USER_ENTERED", body={"values":RN}).execute()
        request = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, 
                            range=("engenharia_de_software!H"+x), valueInputOption="USER_ENTERED", body={"values":([[str(naf)]])}).execute()
        continue

    #Verification if the Student's average is higher or equal to 50 and below 70
    elif (average >= 50) & (average < 70):
      
        naf = 100 - average
        request = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, 
                            range=("engenharia_de_software!G"+x), valueInputOption="USER_ENTERED", body={"values":EF}).execute()
        request = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, 
                            range=("engenharia_de_software!H"+x), valueInputOption="USER_ENTERED", body={"values":([[str(naf)]])}).execute()
        continue
        
















