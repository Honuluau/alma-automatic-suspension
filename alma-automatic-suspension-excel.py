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

print(initials)

def get_file_path():
    return askopenfilename(title="Select Fulfillment - Loans Returns and Overdue Dashboard", filetypes=[("Fulfillment Report", "*.xlsx")])

file_path = get_file_path()

df = pd.read_excel(file_path, header=None)
df_data = df.iloc[13:]

for index, row in df_data.iterrows():
    print(f"Row {index}: {row.tolist()}")

input("Test")