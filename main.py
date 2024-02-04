import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of the spreadsheet.
SAMPLE_SPREADSHEET_ID = "15xT2OWDialERrVxO_pyJUUVdChIxCp-XcufZL9MrBnA"
SAMPLE_RANGE_NAME = "engenharia_de_software!C4:F"


def main():
    # Verify user authentication

    creds = None

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

    print('user authenticated')

    try:
        service = build("sheets", "v4", credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()

        # Read the cells on Google Sheets
        result = (
            sheet.values()
            .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME)
            .execute()
        )

        # Get the values of the cells
        values = result['values']

        # Defining variables

        # Semester absence limit
        absence_limit = 60 * 0.25

        # Values to add
        add_values = []

        # Calculate each student absence, average, status and grade for final test

        for row in values:

            # Student absence
            absence = int(row[0])

            # Student grades
            p1 = float(row[1]) / 10
            p2 = float(row[2]) / 10
            p3 = float(row[3]) / 10

            # Student average
            m = (p1 + p2 + p3) / 3

            # Calculate student status
            if absence > absence_limit:
                # Student status
                sts = 'Reprovado por Falta'
            elif m < 5:
                sts = 'Reprovado por Nota'
            elif 5 <= m < 7:
                sts = 'Exame Final'
            else:
                sts = 'Aprovado'

            if sts == 'Exame Final':
                naf = "{:,.1f}".format(10 - m)
            else:
                naf = 0

            # Verify if the code is running
            print(
                f'Faltas: {absence} / P1: {p1} / P2: {p2} / P3: {p3} / Média: {"{:,.1f}".format(m)} / Situação: {sts}' +
                f' / Nota para Aprovação Final: {naf}')

            # Append values to the variable 'add_values'
            add_values.append([sts, naf])

        # Add the values into spreadsheet

        result = (
            sheet.values()
            .update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range='engenharia_de_software!G4',
                    valueInputOption='USER_ENTERED',
                    body={'values': add_values})
            .execute()
        )

        # Print the spreadsheet is updated
        print('Planilha atualizada.')

    except HttpError as err:
        print(err)


if __name__ == "__main__":
    main()
