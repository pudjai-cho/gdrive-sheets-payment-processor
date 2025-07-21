from google.oauth2 import service_account
from googleapiclient.discovery import build
from pathlib import Path
import re
import pandas as pd
import logging
from dotenv import load_dotenv
import os

logging.basicConfig(level=logging.WARNING,
                    format='%(asctime)s - %(levelname)s - %(message)s')



cur_dir = Path.cwd()
service_account_file = cur_dir / 'service_account.json'
next_month_env_file = cur_dir / '.env.nextmonth'


# load_dotenv()
# payment_column_start = "W"
# payment_column_end = "X"

load_dotenv(dotenv_path=next_month_env_file)
payment_column_start = "X"
payment_column_end = "Y"


SCOPES = ['https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/spreadsheets']

creds = service_account.Credentials.from_service_account_file(
    service_account_file, scopes=SCOPES)

drive_service = build('drive', 'v3', credentials=creds)
sheet_service = build('sheets', 'v4', credentials=creds)

# sheet_name1 = 'OCT-DEC.24'
# # sheet_name2 = 'RENT2025'
# spreadsheet_id = os.getenv('SPREADSHEET_ID')
# sheet_ranges1 = [f"{sheet_name1}!D2:D230", f"{sheet_name1}!O2:Q230"]
# # sheet_ranges2 = [f"{sheet_name2}!D2:D10", f"{sheet_name2}!O2:S10"]



last_row = 202

sheet_name1 = '2025_Rental'
sheet_name2 = ""
spreadsheet_id = os.getenv('SPREADSHEET_ID')
sheet_ranges1 = [f"{sheet_name1}!D2:D{last_row}", f"{sheet_name1}!{payment_column_start}2:{payment_column_end}{last_row}"]


folder_id = os.getenv('FOLDER_ID')
updated_folder_id = os.getenv('UPDATED_FOLDER_ID')
failed_folder_id = os.getenv('FAILED_FOLDER_ID')
filter_pattern = r"\d+[-/]\d+"
row_num1 = 2



if sheet_name2:
    sheet_ranges2 = [f"{sheet_name2}!B2:B15", f"{sheet_name2}!C2:E15"]
    row_num2 = 3
else:
    print("Only one sheet")
    
    
    
sheet = sheet_service.spreadsheets()


def read_sheet(input_sheet_range):
    try:
        sheet_result = sheet.values().batchGet(
            spreadsheetId=spreadsheet_id,
            ranges=input_sheet_range
        ).execute()
        if not sheet_result:
            logging.error("Empty sheet result returned")
            raise ValueError("No data received from sheet")
        return sheet_result
    except Exception as e:
        logging.error(f'Error fetching sheet values: {e}')
        raise e


def align_column_lengths(base_column, editable_column):
    if len(editable_column) < len(base_column):
        col_to_add = len(base_column) - len(editable_column)
        editable_column.extend([[] for _ in range(col_to_add)])


sheet_result = read_sheet(sheet_ranges1)


col_val = sheet_result.get('valueRanges', [])


if not col_val or len(col_val) < 2:
    raise Exception(
        "col_val Expected 2 elements in 'valueRanges', but got fewer or empty list.")
col_room_num_1 = col_val[0].get('values', [])
col_payment_data_1 = col_val[1].get('values', [])
align_column_lengths(col_room_num_1, col_payment_data_1)
col_payment_1 = []
for i, row in enumerate(col_payment_data_1):
    col_payment_1.append(row + [""] * max(0, (row_num1 - len(row))))

if sheet_name2:
    sheet_result2 = read_sheet(sheet_ranges2)
    col_val2 = sheet_result2.get('valueRanges', [])
    if not col_val2 or len(col_val2) < 2:
        raise Exception(
            "col_val2 Expected 2 elements in 'valueRanges', but got fewer or empty list.")

    col_room_num_2 = col_val2[0].get('values', [])
    col_payment_2 = col_val2[1].get('values', [])
    align_column_lengths(col_room_num_2, col_payment_2)
    for i, row in enumerate(col_payment_2):
        col_payment_2[i] = row + [""] * max(0, (row_num2 - (len(row))))


def list_files_in_folders(folder_id):
    query = f"'{folder_id}' in parents and mimeType != 'application/vnd.google-apps.folder' and mimeType != 'application/vnd.google-apps.shortcut'"
    try:
        results = drive_service.files().list(
            q=query,
            fields="files(id, name, webViewLink)"
        ).execute()
        if results.get('files', []):
            return results.get('files', [])
        else:
            logging.error("results.get('files', []) is empty")
            return []

    except Exception as e:
        logging.error(f'Error listing files: {e}')
        return []


def get_room_payment_slips(files):
    df = pd.DataFrame(files)
    df['room'] = df['name'].apply(filter_room_num)
    room_to_link = {}
    room_to_file_id = {}
    for index, row in df.iterrows():
        if row['room']:
            room_number = row['room']
            room_to_link[room_number] = row['webViewLink']
            room_to_file_id[room_number] = row['id']
        else:
            move_file_to_folder(row['id'], failed_folder_id, folder_id)
            print(f"moved {row['name']}")
    return room_to_link, room_to_file_id


def move_file_to_folder(file_id_to_move, target_folder_id, previous_parent):
    try:
        drive_service.files().update(
            fileId=file_id_to_move,
            addParents=target_folder_id,
            removeParents=previous_parent,
        ).execute()
    except Exception as e:
        logging.error(f'Error moving file {file_id_to_move}: {e}')
        raise


def update_sheet_with_hyperlink(row_index, col_index, url, update_sheet_name):
    update_range = f'{update_sheet_name}!{
        chr(ord(payment_column_start) + col_index)}{row_index + 2}'
    formula = f'=HYPERLINK("{url}", "V")'
    body = {"values": [[formula]]}
    try:
        sheet.values().update(
            spreadsheetId=spreadsheet_id,
            range=update_range,
            valueInputOption="USER_ENTERED",
            body=body
        ).execute()
    except Exception as e:
        print(f'Error updating sheet : {e}')


def filter_room_num(input_number: str, pattern_1=r'\(\d+[-/]\d+\)', pattern_2=r'\d+[-/]\d+'):
    if not isinstance(input_number, str):
        logging.error(f"filter_room_num: input must be a string, got {
                      type(input_number)}")
        return None
    matched = re.search(pattern_1, input_number)
    if matched:
        return matched.group()[1:-1].replace("-", "/")
    matched = re.search(pattern_2, input_number)
    if matched:
        return matched.group().replace("-", "/")
    return None


files = list_files_in_folders(folder_id)

if not files:
    logging.error("Slip Folder is empty")
    raise SystemExit

room_to_link, room_to_file_id = get_room_payment_slips(files)
col_room_num_1 = [item for sublist in col_room_num_1 for item in sublist]
col_room_num_1 = list(map(lambda x: filter_room_num(
    x) if isinstance(x, str) else None, col_room_num_1))

if sheet_name2:
    col_room_num_2 = [item for sublist in col_room_num_2 for item in sublist]
    col_room_num_2 = list(
        map(lambda x: filter_room_num(x) if isinstance(x, str) else None, col_room_num_2))

updated_room = []
to_update = []
failed_room = []
full_room = []
updated_room_sheet_2 = []

col_room_num_1_set = set(col_room_num_1)



for key in room_to_link:
    if key in col_room_num_1_set:
        to_update.append(key)
    else:
        failed_room.append(key)


for row_index, room_row in enumerate(col_room_num_1):
    room_number = room_row if room_row else None
    payment_row = col_payment_1[row_index] if row_index < len(
        col_payment_1) else []
    if room_number and room_number in room_to_link:
        print(f"Checking for {room_number}")
        payment_url = room_to_link[room_number]
        print(f"Payment row of {room_number} is {payment_row}")
        if len(col_payment_data_1[row_index]) < row_num1:
            for col_index, cell in enumerate(payment_row):
                if not cell:
                    update_sheet_with_hyperlink(
                        row_index, col_index, payment_url, sheet_name1)
                    print(f"Updated {room_number}")
                    move_file_to_folder(
                        room_to_file_id[room_number], updated_folder_id, folder_id)
                    print(f"Moved {room_number}")
                    updated_room.append(room_number)
                    break
                else:
                    print(f"cell {col_index + 1} of {room_number} is full")
        else:
            full_room.append(room_number)
            print(f"Payment row of {room_number} is full")

if sheet_name2:
    for row_index2, room_row2 in enumerate(col_room_num_2):
        room_number2 = room_row2 if room_row2 else None
        payment_row2 = col_payment_2[row_index2] if row_index2 < len(
            col_payment_2) else []
        if room_number2 and room_number2 in full_room:
            print(
                f"Checking for {room_number2} in second sheet")
            payment_url2 = room_to_link[room_number2]
            print(
                f"Payment row of {room_number2} in second sheet is {payment_row2}")
            for col_index2, cell2 in enumerate(payment_row2):
                if not cell2:
                    update_sheet_with_hyperlink(
                        row_index2, col_index2, payment_url2, sheet_name2)
                    print(
                        f"Updated {room_number2} in second sheet")
                    move_file_to_folder(
                        room_to_file_id[room_number2], updated_folder_id, folder_id)
                    print(f"Moved {room_number2}")
                    updated_room_sheet_2.append(room_number2)
                    break

print(f"Total {len(to_update)} will be updated")
print(to_update)
print(f"Total {len(failed_room)} wasn't matched")
print(failed_room)
print(f"Total {len(full_room)} is full")
print(full_room)
print(f"Total {len(updated_room)} has been updated to Sheet 1")
print(updated_room)
print(f"Total {len(updated_room_sheet_2)} has been updated to Sheet 2")
print(updated_room_sheet_2)
