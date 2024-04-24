from google.oauth2 import service_account
from googleapiclient.discovery import build

spreadsheet_id = "1NE65dbmhbJuM_z1qb2yHiGiguiPsQCEJEhCE0V_5mPQ"
scopes =["https://www.googleapis.com/auth/spreadsheets","https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = service_account.Credentials.from_service_account_file("credentials.json", scopes=scopes)
service = build("sheets", "v4", credentials=credentials)

request = service.spreadsheets().get(spreadsheetId=spreadsheet_id, fields='sheets(data/rowData/values/userEnteredValue,properties(index,sheetId,title))')
sheet_props = request.execute()

total_responses = len(sheet_props['sheets'][0]['data'][0]['rowData'])
print(total_responses)
rowNo=total_responses-1
data=[]
for colNo in range(2,8):
   entry = sheet_props['sheets'][0]['data'][0]['rowData'][rowNo]['values'][colNo]['userEnteredValue']
   for val in entry.values():
      data.append(val)
print(data)
mainBrand = data[0]
noOfComp = data[1]
email = data[2]
channel = data[3].split("=")[-1]
video = data[4].split("=")[-1]
logo = data[5].split("=")[-1]
