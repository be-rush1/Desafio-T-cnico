import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1ZSq1IxwtfEb_Ed4tUBKXMytply-uTZvKEOMlFrnH9uU"

total_classes = 60 #constant for number of classes in a semester

def write_values():
  creds = None
  # Creates token.json file in case it doesn't exist in current directory
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
        "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # keeps credentials in json file
    with open("token.json", "w") as token:
      token.write(creds.to_json())
  try:
    service = build("sheets", "v4", credentials=creds)

    # http requisition to api for reading
    sheet = service.spreadsheets()
    result = (
        sheet.values()
        .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="A4:F27")
        .execute()
    )
  except HttpError as err:
    print(err)

  values = result.get("values", []) #keep values in values matrix

  if not values:
    print("No data found") #check if there are data
    return

  res = [] #matrix to keep average and final situation for each student
  for rows in values:   #computes average grade for each studente
    avg = round((float(rows[3]) + float(rows[4]) + float(rows[5])) / 3)
    absence = (float(rows[2])/total_classes) * 100
    if avg >= 70 and absence <= 25:
      res.append(["Aprovado",0])

    if avg >= 50 and avg < 70 and absence <= 25:
      res.append(['Exame Final', 100-avg])

    if avg < 50:
      res.append(['Reprovado por nota', 0])

    if absence > 25:
      res.append(['Reprovado por falta', 0])

  try:
    # http requisition for writing
    service = build("sheets", "v4", credentials=creds)
    body = {"values": res} #parse res as body for write requisition
    result = (
      service.spreadsheets()
        .values()
        .update(
        spreadsheetId=SAMPLE_SPREADSHEET_ID,
        range="G4:H27", #write from positon G4 to H27
        valueInputOption="USER_ENTERED", #keep data type on spreead sheet
        body=body,
        )
        .execute()
    )
  except HttpError as error:
    print(f"An error occurred: {error}")
    return error

  print(f"{result.get('updatedCells')} cells updated.") #prints number of cells

if __name__ == "__main__":

 write_values()