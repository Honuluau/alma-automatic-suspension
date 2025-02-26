import openpyxl
import time
import sys
import os
import operator
from datetime import date

url = ""
initials = ""

# This is a loop to confirm your initials to display near the end of the suspension note.
while (initials == ""):
    newInitials = input("Please insert your initials: ")
    confirmation = input("The format for the note will be: -" + str(newInitials) + "  | Is this correct? Y or N \n")

    if str.lower(confirmation) == "y":
        initials = newInitials

# This is a loop to get a working directory url to the excel sheet. If the format has backslashes, it will be convertered for use.
while (url == ""):
    newUrl = input("Please drag the file to this window, and make sure this window is selected, then press enter:\n")
    newUrl = newUrl.replace("\"", "")
    newUrl = newUrl.replace("\\", "/")

    print("Filepath recieved: " + newUrl)
    if os.path.exists(newUrl):
        url = newUrl
    else:
        print("File not found.")

# OpenPyXL is a python library to read/write Excel files.
# It runs locally and makes know pulls or requests to send data through the internet and is completely safe to use.
workbook = openpyxl.load_workbook(url)
activeSheet = workbook.active

data = {}

'''
data structure
data
    - userId 
        -- row
        -- days overdue
        -- item#
            --- item title 
            --- process status 
'''
previousId = None # placeholder
previousIterator = 0

def addData(id):
    global previousId
    global previousIterator
    # new user id.
    data[id] = {}
    data[id]["Items"] = {}
    data[id]["Row"] = row
    data[id]["DaysOverdue"] = activeSheet["G"][row].value
    data[id]["Name"] = activeSheet["C"][row].value + ", " + activeSheet["B"][row].value
    data[id]["Items"]["Item1"] = {
        "Title": activeSheet["I"][row].value,
        "ProcessStatus": activeSheet["J"][row].value,
        "Barcode": activeSheet["H"][row].value
    }

    previousId = id
    previousIterator = 1
    print("[AAS] New UserId Found: " + str(id))


# This section organizes it into data.
for row in range(0, activeSheet.max_row):

    # Iterates over sheet to find the user id rows and to not cause dupliates with multiple lost items.
    for col in activeSheet.iter_cols(1,1):
        id = col[row].value

        if id == None:
            if previousId != None:
                previousIterator += 1
                print("[AAS] Another Item Found, Item " + str(previousIterator))

                daysOverdue = activeSheet["G"][row].value
                if data[previousId]["DaysOverdue"] < daysOverdue:
                       data[previousId]["DaysOverdue"] = daysOverdue

                data[previousId]["Items"]["Item" + str(previousIterator)] = {
                "Title": activeSheet["I"][row].value,
                "ProcessStatus": activeSheet["J"][row].value,
                "Barcode": activeSheet["H"][row].value
            }
            continue
        if isinstance(id, str):
            if id.isnumeric():
                addData(id)
            else:
                if id[1:].isnumeric():
                    addData(id)
        else:
            addData(id)


sorted_data = dict(sorted(data.items(), key=lambda x: x[1]['DaysOverdue']))
currentDate = date.today()
outputFilePath = os.path.expanduser("~") + "/alma-automatic-suspension-output-" + str(currentDate.month) + "-" + str(currentDate.day) + "-" + str(currentDate.year) + ".txt"
breakLine = False
legalLetterRequirement = 30
with open(outputFilePath, "w", encoding='utf-8') as file:
    for id in sorted_data:
        # This is a person.

        suspensionNote = "SUSPENDED / Instance#X / LOST ["
        itemsLost = ""
        itemBarcodes = ""
        for item in data[id]["Items"]:
            if data[id]["Items"][item]["ProcessStatus"] == "LOST":
                        itemsLost = itemsLost + "\'" + data[id]["Items"][item]["Title"] + "\',"
                        itemBarcodes = itemBarcodes + str(data[id]["Items"][item]["Barcode"]) + ","

        if itemsLost == "":
            # This person does not match the requirements to be suspended, their items are overdue but not lost.
            continue
        else:
            itemsLost = itemsLost[:len(itemsLost)-1]
            itemBarcodes = itemBarcodes[:len(itemBarcodes)-1]
            suspensionNote = suspensionNote + itemsLost + "]-unresolved- " + str(currentDate.month) + "/" + str(currentDate.day) + "/" + str(currentDate.year) + "-" + initials
            if data[id]["DaysOverdue"] > legalLetterRequirement and breakLine == False:
                breakLine = True
                file.write("\n\n\n\nThe following are people who are eligible to receive a legal letter. Fees are not taken into an account int this list. Requirement: " + str(legalLetterRequirement) + " days overdue.\n\n\n\n")
                 

            file.write("UserId: " + str(id) + "\nName: " + str(data[id]["Name"]) + "\n" + str(data[id]["DaysOverdue"]) + " days overdue.\n" + suspensionNote + "\nThese are the item barcodes in order: " + itemBarcodes + "\n")

try:
    os.system(outputFilePath)
except Exception as e:
    print("An error occured while opening the notepad, but can still be found under the Documents folder.")

print("This window will close in 30 seconds.")
time.sleep(30)