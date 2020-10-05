import pickle
import os.path
import os
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from datetime import datetime
import string
import configparser

config = configparser.ConfigParser()
cwd = os.path.dirname(os.path.abspath(__file__))
config.read(os.path.join(cwd, "config.ini"))

alphabet = list(string.ascii_uppercase)


def generate_columns(starting_column):
    columns = []
    index_starting_column = 0
    for index, item in enumerate(alphabet):
        if item == starting_column:
            index_starting_column = index
            break

    for i in range(index_starting_column, len(alphabet)):
        columns.append(alphabet[i])
        if i == len(alphabet) - 1:
            for letter in alphabet:
                for second_letter in alphabet:
                    columns.append(f"{letter}{second_letter}")

    return columns


SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

sections = config.sections()
SAMPLE_SPREADSHEET_ID_input = config["default"]["spreadsheet_id"]
SHEET_NAME = config["default"]["sheet_name"]

row_gf = "11"
row_me = "30"

column = "A"  # beginning of column
COLUMNS = generate_columns(starting_column=column)
print(COLUMNS)

# each day, our configuration file will be updated
# we need to check if we try to launch the script in the same day
# in this case, there is no need to update anything
new_day = False
if config["default"]["date"] == datetime.today().strftime("%Y-%m-%d"):
    print("yoh! same day bro, same day...")
else:
    new_day = True
# we now have what we need: the starting column and the row number
# now it's time to check the corresponding column for a given date
# as I wrote those lines, we are 2020-07-19 and the corresponding column is "U" for this month
# I will work on something later so when we change month, we won't need to update the starting values manually
# ok, let's go! we will store the current date, with the column, in a file
if new_day:
    previous_column = config["default"]["corresponding_column"]
    for i, j in enumerate(COLUMNS):
        if j == previous_column:
            column = COLUMNS[i + 1]
            break

    config_file = open("config.ini", "w+")
    config.set("default", "date", datetime.today().strftime("%Y-%m-%d"))
    config.set("default", "corresponding_column", column)
    config.write(config_file)
    config_file.close()


def main():
    creds = None

    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)

    service = build("sheets", "v4", credentials=creds)

    attendance = ""
    if datetime.today().weekday() == 5 or datetime.today().weekday() == 6:
        print("it's the weekend, what the hell are you doing?")
        return
    else:
        attendance = "1"
    values = [
        [
            # Cell values ...
            attendance
        ],
        # Additional rows ...
    ]
    body = {"values": values}
    sheet = service.spreadsheets()

    # retrieving all our names in the "B" column
    result = (
        sheet.values()
        .get(spreadsheetId=SAMPLE_SPREADSHEET_ID_input, range=f"{SHEET_NAME}!B1:B40")
        .execute()
    )
    names = result.get("values", [])
    for i, name in enumerate(names):
        if name:
            if name[0].find("Edwige") > 0:
                row_gf = i + 1
            if name[0].find("Tefy") > 0:
                row_me = i + 1

    sheet.values().update(
        spreadsheetId=SAMPLE_SPREADSHEET_ID_input,
        range=f"{SHEET_NAME}!{column}{row_gf}:{column}{row_gf}",
        valueInputOption="USER_ENTERED",
        body=body,
    ).execute()
    sheet.values().update(
        spreadsheetId=SAMPLE_SPREADSHEET_ID_input,
        range=f"{SHEET_NAME}!{column}{row_me}:{column}{row_me}",
        valueInputOption="USER_ENTERED",
        body=body,
    ).execute()


if __name__ == "__main__":
    main()
