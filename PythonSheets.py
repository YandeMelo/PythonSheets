import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import logging

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Set the log level as needed (DEBUG, INFO, WARNING, ERROR, CRITICAL)
logging.basicConfig(level=logging.INFO)

def authenticate_google_sheets():
    # Authenticate and authorize access to Google Sheets API.
    if os.path.exists("token.json"):
        credentials = Credentials.from_authorized_user_file("token.json", SCOPES)

    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("client_secret.json", SCOPES)
            credentials = flow.run_local_server(port=0)

        with open("token.json", "w") as token:
            token.write(credentials.to_json())

    return credentials

def calculate_passing_grade(average, situation):
    # Calculate passing grade based on the situation.
    if situation == "Exame Final":
        return round(100 - average)
    else:
        return 0

def main():
    # Main function to update Google Sheets with evaluation results.
    credentials = None
    credentials = authenticate_google_sheets()

    try:
        service = build("sheets", "v4", credentials=credentials)
        sheet = service.spreadsheets()
        
        add_values = []
        # List to store information and update in the spreadsheet at the end
        result = (
            sheet.values()
            .get(spreadsheetId="1nyEeA-tUxuHUquDWbo3iH8ZY1Jh3xlDV4WSgZ5mDZZk", range="engenharia_de_software!A:A")
            .execute()
        )
        total_rows = result.get("values", [])
        
        # Check if there is at least one row in the spreadsheet
        if total_rows:
            total_rows = len(total_rows)
            
            # Use all available rows
            for row in range(4, total_rows + 1):
              result_score = (
                  sheet.values()
                  .get(spreadsheetId="1nyEeA-tUxuHUquDWbo3iH8ZY1Jh3xlDV4WSgZ5mDZZk", range=f"engenharia_de_software!D{row}:F{row}")
                  .execute()
              )
              result_absence = (
                  sheet.values()
                  .get(spreadsheetId="1nyEeA-tUxuHUquDWbo3iH8ZY1Jh3xlDV4WSgZ5mDZZk", range=f"engenharia_de_software!C{row}")
                  .execute()
              )
              
              absence = int(result_absence.get("values", [["0"]])[0][0])
              # Get the absence values present in the spreadsheet
              
              score = result_score.get("values", [["", "", ""]])[0]
              score_float = [float(value) if value else 0 for value in score]
              # Get the grades' values present in the spreadsheet and transform them into FLOAT
              
              average = sum(score_float) / len(score_float)
              # Calculate the average
              
              # Log the average
              logging.info(f"Average for row {row}: {average}")
              
              # Evaluate the situation and calculate passing grade
              if absence > 15:
                  situation = "Reprovado por Falta"
              elif average < 50:
                  situation = "Reprovado por Nota"
              elif 50 <= average < 70:
                  situation = "Exame Final"
              else:
                  situation = "Aprovado"
              
              passing_grade = calculate_passing_grade(average, situation)
              add_values.append([situation, passing_grade])
        else:
            logging.warning("The spreadsheet is empty.")
            
        # Log the values to be added
        logging.info(f"Values to be added: {add_values}")
        
        update_range = f"engenharia_de_software!G4:H{total_rows + 3}"
        result = (
            sheet.values()
            .update(
                spreadsheetId="1nyEeA-tUxuHUquDWbo3iH8ZY1Jh3xlDV4WSgZ5mDZZk",
                range=update_range,
                valueInputOption="USER_ENTERED",
                body={"values": add_values},
            )
            .execute()
        )
        
        # Log the successful update
        logging.info("Update successful.")
    
    except HttpError as err:
        # Log the HTTP error
        logging.error(f"HTTP error: {err}")

if __name__ == "__main__":
    main()
