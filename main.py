import copy
import json
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2 import service_account

from copy import deepcopy
from pynput import keyboard

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly", "https://www.googleapis.com/auth/spreadsheets"]

with open("secrets.json", "r") as secrets_file:
    secrets = json.load(secrets_file)
    SPREADSHEET_ID = secrets["SPREADSHEET_ID"]
    API_KEY = secrets["API_KEY"]
secrets_file.close()


# global mondstadt_sheet,liyue_sheet,inazuma_sheet,sumeru_sheet,fontaine_sheet,inventory_sheet,overall_sheet
region_sheet_mapping = {0: "Mondstadt", 1: "Liyue", 2: "Inazuma", 3: "Sumeru", 4: "Fontaine", 5: "Natlan"}
bossmat_mapping = {}


def do_setup():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
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
            os.remove("token.json")
            creds.token = None
            creds.refresh(Request())

        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    service = build("sheets", "v4", credentials=creds)
    return service.spreadsheets()


def main():
    global sheet, books_total, bossmats_total
    sheet = do_setup()
    mondstadt_sheet = (sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Mondstadt_chars").execute())
    liyue_sheet = (sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Liyue_chars").execute())
    inazuma_sheet = (sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Inazuma_chars").execute())
    sumeru_sheet = (sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Sumeru_chars").execute())
    fontaine_sheet = (sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Fontaine_chars").execute())
    natlan_sheet = (sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Natlan_chars").execute())
    inventory_sheet = (sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Item Inventory!A1:P13").execute())
    overall_sheet = (sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Overall!A1:P13").execute())
    sheets = [mondstadt_sheet, liyue_sheet, inazuma_sheet, sumeru_sheet, fontaine_sheet, natlan_sheet, inventory_sheet, overall_sheet]

    initialize_bosses()
    initialize_books()

    for i in range(len(region_sheet_mapping)):
        calculate_region(sheets[i].get("values", []), i)

    books_total = copy.deepcopy(books_needed)
    bossmats_total = copy.deepcopy(bossmats_needed)

    do_diff()
    #calculate_holdovers()
    print(bossmats_needed)
    print(books_needed)
    #print(bossmat_mapping)
    #record(overall_sheet)
    record_books()
    record_boss()
    return

    # sheets = [mondstadt_sheet, liyue_sheet, inazuma_sheet, sumeru_sheet, fontaine_sheet]
    try:
        result = (sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Item Inventory!A1:P13").execute())
        values = result.get("values", [])

        for row in values:
            print(row)

    except HttpError as err:
        print(err)
    # initialize_materials(sheet)


bossmats_needed = {}
bossmats = {}
books_needed = {}
books = {}
books_total = {}
bossmats_total = {}

def get_boss(level):
    level = level+1
    if level in [7,8]:
        return 1
    elif level in [9,10]:
        return 2
    else:
        return 0


def get_rarity(level):
    level = level + 1
    if level in [3,4,5,6]: return 1
    elif level in [7,8,9,10]: return 2
    else: return 0


def initialize_bosses():
    global bossmats, bossmat_mapping
    print("Initializing Boss Inventory")
    data = (sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Inventory_Bossmats").execute())
    values = data.get("values", [])

    #Basically converts numbers to ints
    values = list(map(lambda x: list(map(lambda y: int(y) if y.isnumeric() else y, x)), values))

    #I Think this is removing the empty spots
    values = list(map(lambda x: list(filter(lambda y: y != '', x)), values))

    #Removing empty lists I suppose
    values = list(filter(lambda x: x != [], values))

    skip = 4
    for i in range(len(values)):
        if skip == 4:
            print(values[i])
            bossmat_names = list(filter(lambda x: x != "", values[i]))
            temp_bossmat_dict = dict.fromkeys(bossmat_names, [])
            temp_bossmat_index = 0
            #print(bossmat_mapping)
            skip = 0
        else:
            for j in range(len(values[i])):
                if isinstance(values[i][j], str):
                    bossmats[values[i][j]] = values[i][j + 1]
                    bossmats_needed[values[i][j]] = 0
                    current_boss = bossmat_names[temp_bossmat_index]
                    current_item = values[i][j]
                    temp_bossmat_dict[current_boss] = temp_bossmat_dict[current_boss] + [current_item]

                    if temp_bossmat_index + 1 < len(temp_bossmat_dict):
                        temp_bossmat_index = temp_bossmat_index + 1
                    if temp_bossmat_index == 3:
                        bossmat_mapping.update(temp_bossmat_dict)
                        temp_bossmat_index = 0
                       # print(temp_bossmat_dict)
        skip = skip + 1

    common = (sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Item Inventory!Crowns_Solvents").execute()).get("values", [])[0]
    bossmats["Crown of Insight"] = int(common[0])
    bossmats["Dream Solvent"] = int(common[1])

    bossmats_needed["Crown of Insight"] = 0
    bossmats_needed["Dream Solvent"] = 0
    print(bossmats)
    print(len(bossmats))


def initialize_books():
    global books, books_needed
    print("Reading books in inventory")
    data = (sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Item Inventory!Inventory_Books").execute())
    values = data.get("values", [])

    for i in range(len(values)):
        values[i] = list(map(lambda x: int(x) if x.isnumeric() else x, values[i]))

    levels = ["Green", "Blue", "Purple"]

    for i in range(len(values)):
        for j in range(len(values[i])):
            temp = values[i][j]
            if values[i][j] not in levels and values[i][j] != '' and isinstance(values[i][j], str):
                books[values[i][j]] = [values[i + 1][j], values[i + 2][j], values[i + 3][j]]
                books_needed[values[i][j]] = [0,0,0]

    print(books)
    print(len(books))


def calculate_region(region_sheet, index):
    global books_needed, bossmats_needed
    print(f"Calculating books for {region_sheet_mapping[index]}")
    for i in range(1, len(region_sheet)):
        if region_sheet[i][0] in books:
            for j in range(0, len(region_sheet[i])):
                if region_sheet[i][j] != "" and region_sheet[i][j] in books:
                    calculate_mats_needed(region_sheet[i][j], region_sheet[i+1][j], [region_sheet[i][j+1], region_sheet[i+1][j+1], region_sheet[i+2][j+1]])
    #print(books_needed, bossmats_needed)


level_mats = [-100000, 3, 2, 4, 6, 9, 4, 6, 12, 16]


def calculate_mats_needed(book, bossmat, levels):
    global books_needed, bossmats_needed, books_total, bossmats_total
    print(f"{book}  {bossmat} {levels}")
    levels = list(map(lambda x: int(x) if x.isnumeric() else x, levels))

    for level in levels:
        for current_level in range(level, 9):
            books_needed[book][get_rarity(current_level)] = books_needed[book][get_rarity(current_level)] + level_mats[current_level]
            print(f"Doing {bossmat}:{current_level} {bossmats_needed[bossmat]} + {get_boss(current_level)}")
            bossmats_needed[bossmat] = bossmats_needed[bossmat] + get_boss(current_level)



def calculate_holdovers():
    global bossmats_needed, bossmats

    for boss_name, boss_mats in bossmat_mapping.items():
        print(f"Calculating holdovers for {boss_name}")
        for boss_mat in boss_mats:
            print(f"{boss_mat}: {bossmats[boss_mat]}/{bossmats_needed[boss_mat]}")
        filtered_bossmats_needed = dict(filter(lambda item: item[0] in boss_mats, bossmats_needed.items()))
        needed = sum(filtered_bossmats_needed.values())

        if needed > 0:
            print(f"Dream Solvents left: {bossmats['Dream Solvent']}")
            filtered_bossmats = dict(filter(lambda item: item[0] in boss_mats, bossmats.items()))
            have = sum(filtered_bossmats.values())

            if have - needed > 0 and needed < bossmats["Dream Solvent"]:
                bossmats["Dream Solvent"] = bossmats["Dream Solvent"] - needed
                print(f"Using {have-needed} solvents?")
                for boss_mat in boss_mats:
                    bossmats_needed[boss_mat] = 0
                print(dict(filter(lambda item: item[0] in boss_mats, bossmats_needed.items())))
            print(f"Dream Solvents after holdover calculation: {bossmats['Dream Solvent']}")
        #print()

    # for book in books.keys():
    #     print(f'{book} {books_needed[book]} {books[book]}')
    #     have = books[book]
    #     needed = books_needed[book]
    #     extra_books = [have - needed for have, needed in zip(have, needed)]
    #     print(extra_books)
    #
    #     if extra_books[1] > 0 and extra_books[0] > 1:
    #         extra_books[1] = extra_books[1] + extra_books[0] // 3
    #         extra_books[0] = extra_books[0] % 3
    #
    #         print(extra_books)
    #         extra_books[2] = extra_books[2] + extra_books[1] // 3
    #         extra_books[1] = extra_books[1] % 3
    #
    #         print(extra_books)


def do_diff():
    global books_needed, bossmats_needed

    for book in books.keys():
        books_needed[book][0] = books_needed[book][0] - books[book][0]
        books_needed[book][1] = books_needed[book][1] - books[book][1]
        books_needed[book][2] = books_needed[book][2] - books[book][2]

    for bossmat in bossmats.keys():
        bossmats_needed[bossmat] = bossmats_needed[bossmat] - bossmats[bossmat]

    # Prevents negatives from being generated
    # for k, v in bossmats_needed.items():
    #     if v < 0:
    #         bossmats_needed[k] = 0

    #Bitch bastard that sets all overlflow books to zero
    # for k in books_needed:
    #     books_needed[k] = list(map(lambda x: x*0 if x < 0 else x, books_needed[k]))


def setup_serivce():
    credentials = service_account.Credentials.from_service_account_file('genshin-material-calculator-dd1ac2710477.json')
    service = build("sheets", "v4", credentials=credentials)
    return service.spreadsheets()


def build_book_printout(needed, total):
    # 3 cases
    # Books needed < total == Normal case, need more books to farm SMALLER/BIGGER
    # Books needed > total, total != 0, Overflow case BIGGER/SMALLER
        # Right now, with how we're recording values, it will be negative
        # I suppose we can just add to total?
    # BN > total, total == 0, Done, record 0
    # BN == total, spreadsheet shows 0
         # Want to show matching number
    #print(f"{needed}    {total}")
    if needed < 0 and total != 0:
        return f"{abs(needed)+total}/{total}"
    elif needed == 0 and total > 0:
        return f"{total}/{total}"
    elif needed > total != 0:  #Books needed > total, total != 0, Overflow case BIGGER/SMALLER <--- I don't think this will happen, always negative
        return f"{needed+total}/{total}"
    elif needed > total == 0:
        return "0"
    elif total == 0:
        return "0"
    else:
        return f"{needed}/{total}"


def record_books():
    print("Writing materials needed")
    data = (sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Materials_Overview",
                               majorDimension="COLUMNS").execute())
    values = data.get("values", [])
    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

    for x in range(len(values))[::4]:
        column_index = x
        column = values[column_index]
        for y in range(len(column)):
            row_index = y
            cell = column[y]
            if cell in books.keys():
                cells_range_start = f"{letters[column_index+1]}{row_index+1}"
                cells_range_end = f"{letters[column_index+3]}{row_index+1}"
                green_index = f"{cells_range_start}"
                blue_index = f"{letters[column_index+2]}{row_index+1}"
                purple_index = f"{cells_range_end}"
                book_row = (sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                                          range=f"{cells_range_start}:{cells_range_end}",
                                                          majorDimension="COLUMNS").execute()).get("values", [])
                print(f"{cell}: {book_row}")
                print(f"{cells_range_start}:{cells_range_end}")

                green_previous =  book_row[0][0]
                blue_previous = book_row[1][0]
                purple_previous = book_row[2][0]

                green_string = build_book_printout(books_needed[cell][0], books_total[cell][0])
                blue_string = build_book_printout(books_needed[cell][1], books_total[cell][1])
                purple_string = build_book_printout(books_needed[cell][2], books_total[cell][2])

                if green_string != green_previous:
                    update_cell(cell_index=green_index, item=cell, value=green_string)
                if blue_string != blue_previous:
                    update_cell(cell_index=blue_index, item=cell, value=blue_string)
                if purple_string != purple_previous:
                    update_cell(cell_index=purple_index, item=cell, value=purple_string)

                print(f"{cell} - New:{green_string} Old:{green_previous}")
                print(f"{cell} - New:{blue_string} Old:{blue_previous}")
                print(f"{cell} - New:{purple_string} Old:{purple_previous}")


def record_boss():
    print("Writing Boss Materials needed")
    data = (sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Materials_Overview",
                               majorDimension="COLUMNS").execute())
    values = data.get("values", [])
    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

    for x in range(len(values))[::4]:
        column_index = x
        column = values[column_index]
        for y in range(len(column)):
            row_index = y
            cell = column[y]
            if cell in bossmats.keys():
                cell_index = f"{letters[column_index + 1]}{row_index + 1}"

                boss_row = (sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                               range=f"{cell_index}",
                                               majorDimension="COLUMNS").execute()).get("values", [])
                print(f"{cell}: {boss_row}")
                print(f"{cell_index}")

                boss_string = "0"
                boss_needed = bossmats_needed[cell]
                boss_total_needed = bossmats_total[cell]

                if boss_total_needed != 0:
                    if boss_needed < 0:
                        boss_string = f"{abs(boss_needed - boss_total_needed)}/{boss_total_needed}"
                    else:
                        boss_string = f"{boss_needed}/{boss_total_needed}"

                boss_previous = boss_row[0][0]

                if boss_string != boss_previous:
                    update_cell(cell_index=cell_index, item=cell, value=boss_string)

                print(f"{cell} - New:{boss_string} Old:{boss_previous}")

def update_cell(cell_index, item, value):
    sheet.values().update(spreadsheetId=SPREADSHEET_ID,
                          key='AIzaSyAGHJyMe1eEevgspPhehzI_mDJ3imwO0Eo',
                          range=f"Overall!{cell_index}",
                          valueInputOption="USER_ENTERED",
                          body={'majorDimension': 'COLUMNS', 'values': [[value]]}).execute()
    print(f"Wrote {item}: {value} in {cell_index}")

# with keyboard.GlobalHotKeys({
#         '-+*': main}) as h:
#         h.join()
main()