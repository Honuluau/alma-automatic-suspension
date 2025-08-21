import pandas as pd
import datetime
from datetime import date
import os
import csv
from tkinter.filedialog import askopenfilename

current_date = date.today()
initials = ""

def set_initials():
    global initials

    while initials == "":
        new_initials = input("Set initials: [SUSPENSION NOTE]-")
        confirmation = input(f"The format for the note will be -{new_initials} | Is this correct? Y or N\n> ")

        if str.lower(confirmation) == "y":
            initials = new_initials

def get_file_path():
    print("Asking for file. Expect pop-up window.")
    file_path = askopenfilename(title="Select Fulfillment - Loans Returns and Overdue Dashboard", filetypes=[("Fulfillment Report", "*.xlsx")])
    print(f"Found: {file_path}")
    return file_path

def get_df_data(file_path):
    df = pd.read_excel(file_path, header=None)
    return df.iloc[13:] # Skips to row 13 because the fulfillment report has a lot of blank space.

def iterate_rows_to_form_data(df_data):
    data = {}

    current_id = 0
    for index, row in df_data.iterrows():
        row_data = row.tolist()
        eagle_id = row_data[0]
        days_overdue = row_data[6]
        recent_due = row_data[5]
        if pd.isna(eagle_id):
            # Adding Item
            data[current_id]["items"].append((row_data[8], row_data[7]))

            # Update overdue day
            if days_overdue > data[current_id]["overdue"]:
                data[current_id]["overdue"] = days_overdue

            # Update recent overdue day
            if recent_due > data[current_id]["recent"]:
                data[current_id]["recent"] = recent_due
        else:
            # Creating new data
            current_id = eagle_id

            data[eagle_id] = {}
            data[eagle_id]["first_name"] = row_data[1]
            data[eagle_id]["last_name"] = row_data[2]
            data[eagle_id]["recent"] = row_data[5]
            data[eagle_id]["overdue"] = days_overdue
            data[eagle_id]["items"] = []

            # title, barcode
            data[eagle_id]["items"].append((row_data[8],row_data[7]))

    sorted_data = dict(sorted(data.items(), key=lambda x: x[1]['overdue']))
    return sorted_data

# Get Alma-Automatic-Suspension Logs folder
def get_aasl_folder():
    output_path = f"{os.path.expanduser("~")}/Documents"
    folder_name = "Alma-Automatic-Suspension Logs"

    try:
        os.makedirs(f"{os.path.expanduser("~")}/Documents/{folder_name}")
        print(f"Created {folder_name} in {output_path}")
        output_path = f"{output_path}/{folder_name}"
    except:
        print(f"Directory found in: {output_path}")
        return f"{output_path}/{folder_name}"

def format_name(name):
    name = name.rstrip(" /")
    return name

def format_suspension_note(items):

    item_string = "[ "
    for item in items:
        item_string += f"(\'{format_name(item[0])}\', {item[1]}), "
    item_string = item_string[:-2] + " ]"

    return f"SUSPENDED / Instance#X / LOST {item_string} -unresolved- {current_date} -{initials}"

def write_csv_log(data, path):
    rows = [["Eagle Id", "Name", "Most Recent Overdue","Longest Overdue (Days)", "Number of Items", "Suspension Note"]]
    for eagle_id in data:
        rows.append([
            eagle_id,
            f"{data[eagle_id]["last_name"]},"
            f" {data[eagle_id]["first_name"]}",
            data[eagle_id]["recent"],
            data[eagle_id]["overdue"],
            len(data[eagle_id]["items"]),
            format_suspension_note(data[eagle_id]["items"])
        ])

    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    file_path = f"{path}/AAS-{timestamp}.csv"
    with open(file_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerows(rows)

    try:
        os.startfile(file_path)
    except Exception as e:
        # This has never happened during testing but is a precaution just in case something goes wrong.
        print("An error occurred while opening the log, you can open it through ~/Documents/Alma-Automatic-Suspension Logs")

path = get_aasl_folder()
data = iterate_rows_to_form_data(get_df_data(get_file_path()))
set_initials()
write_csv_log(data, path)