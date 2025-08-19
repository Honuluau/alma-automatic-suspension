import pandas as pd
from datetime import *
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
    return askopenfilename(title="Select Fulfillment - Loans Returns and Overdue Dashboard", filetypes=[("Fulfillment Report", "*.xlsx")])

def get_df_data(file_path):
    df = pd.read_excel(file_path, header=None)
    return df.iloc[13:] # Skips to row 13 because the fulfillment report has a lot of blank space.

def iterate_rows_to_form_data(df_data):
    data = {}

    current_id = 0
    for index, row in df_data.iterrows():
        row_data = row.tolist()
        eagle_id = row_data[0]

        if pd.isna(eagle_id):
            # Adding Item
            print("Adding item.")
        else:
            # Creating new data
            current_id = eagle_id

            data[eagle_id] = {}
            data[eagle_id]["first_name"] = row_data[1]
            data[eagle_id]["last_name"] = row_data[2]
            data[eagle_id]["items"] = {}

            print(f"Adding user data for {eagle_id} ({row_data[1], row_data[2]})")

    input("Test")

iterate_rows_to_form_data(get_df_data(get_file_path()))