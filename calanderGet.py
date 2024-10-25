import datetime
import os.path

import google.auth
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/calendar.readonly","https://www.googleapis.com/auth/spreadsheets"]
startList = []
endList = []
summaryList = []
spreadSheetID = "17ytMl_Z9XLYz1EJpNJ-rSaygH01i4ejUnSsLYWDHeNM"
def calanderGetting (service):
      # Call the Calendar API
    now = datetime.datetime.now().isoformat() + "Z"  # 'Z' indicates UTC time
    print("Getting the upcoming 10 events")
    events_result = (
        service.events()
        .list(
            calendarId="primary",
            timeMin=now,
            maxResults=10,
            singleEvents=True,
            orderBy="startTime",
        )
        .execute()
    )
    events = events_result.get("items", [])

    if not events:
      print("No upcoming events found.")
      return

    # Prints the start and name of the next 10 events
    for event in events:
    
      start = event["start"].get("dateTime", event["start"].get("date"))
      startList.append(start)
      end = event["end"].get("dateTime", event["end"].get("date"))
      endList.append(end)
      print(start, end, event["summary"])
      summaryList.append(event["summary"])
    return startList, endList, summaryList      
  
def sheetsWrite(service, spreadsheetID, rangeName, valueInputOption, _values):
  # #print(_values+ "len : " +len(_values)+ " len 2 : "+ len(_values[0]))
  # print("len")
  # print(len(_values))
  # print("len2 : ")
  # print(_values[0])  
  values = [
    [_values[0],_values[2]],[_values[1]],
]
  values2 = [[] for i in range(len(_values[0]))]
  for row in range(0, len(_values[0])):
    for column in range(0, len(_values)):
      print("row"+str(row)+ str(len(_values[0])))
      values2[row].append(_values[column][row])
  print(values2)
  data = [
    {"range": rangeName, "values": values2},
          # Additional ranges to update ...
  ]
  body = {"valueInputOption": valueInputOption, "data": data}
  result = (
    service.spreadsheets()
    .values()
    .batchUpdate(spreadsheetId=spreadsheetID, body=body)
    .execute()
    )
  print(f"{(result.get('totalUpdatedCells'))} cells updated.")
  return result


def main():

  """Shows basic usage of the Google Calendar API.
  Prints the start and name of the next 10 events on the user's calendar.
  """
  creds = None

  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
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
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    calanderService = build("calendar", "v3", credentials=creds)
    sheetsService = build("sheets", "v4", credentials=creds)
    
    sheetsWrite(sheetsService, spreadSheetID, "A1:G10","USER_ENTERED",calanderGetting(calanderService))
    
    
  except HttpError as error:
    print(f"An error occurred: {error}")





if __name__ == "__main__":
  main()