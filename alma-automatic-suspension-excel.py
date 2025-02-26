import openpyxl
import time
import sys
import os
import operator
from datetime import date

# The date.
currentDate = date.today()

# Directory URL which will be assigned from user-input.
# Initials to go at the end of the suspension note, also assigned from user-input.
# outputFilePath refers to the user's home, on windows it is the "Documents" folder and creates a new text file named after the date.
url = ""
initials = ""
outputFilePath = os.path.expanduser("~") + "/alma-automatic-suspension-output-" + str(currentDate.month) + "-" + str(currentDate.day) + "-" + str(currentDate.year) + ".txt"

# These variables are in regards to legal letters. BreakLine is the line to seperate the legal letter suspensions.
# legalLetterRequirement is the count of days an item can be overdue before being sent a legal letter, if it's one day over it gets seperated.
breakLine = False
legalLetterRequirement = 30

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
# It is a little bit out of date but can work for our purpose since the data is also in an old format. (2010 Excel Files)
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

# Adds to the dictionary where the key is the userId, and the values are another dictionary. 
def addData(id):
    global previousId
    global previousIterator
    # Assigning a new user Id.
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

    # Assigning previous variables to find multiple items in the sheet.
    previousId = id
    previousIterator = 1
    print("[AAS] New UserId Found: " + str(id))


# This section organizes it into data.
for row in range(0, activeSheet.max_row):

    # Iterates over sheet to find the user id rows and to not cause dupliates with multiple lost items.
    for col in activeSheet.iter_cols(1,1):
        id = col[row].value

        # A blank entry was listed which in most cases is another item, this assigns new items to the previous id.
        if id == None:
            if previousId != None:
                previousIterator += 1
                print("[AAS] Another Item Found, Item " + str(previousIterator))

                # This if statement checks to see if the new item is more recently overdue to sort the id higher later on the list.
                daysOverdue = activeSheet["G"][row].value
                if data[previousId]["DaysOverdue"] < daysOverdue:
                       data[previousId]["DaysOverdue"] = daysOverdue

                data[previousId]["Items"]["Item" + str(previousIterator)] = {
                "Title": activeSheet["I"][row].value,
                "ProcessStatus": activeSheet["J"][row].value,
                "Barcode": activeSheet["H"][row].value
            }
            continue
        # This checks to see if the id is a string since Community Member's id's start with a letter.
        if isinstance(id, str):
            if id.isnumeric():
                addData(id)
            else:
                # Community Member Id
                if id[1:].isnumeric():
                    addData(id)
        else:
            addData(id)

# Sorts data using lambda to the DaysOverdue from earliest to longest overdues.
sorted_data = dict(sorted(data.items(), key=lambda x: x[1]['DaysOverdue']))

# Opens the output file path to start writing in the utf-8 encoding format.
with open(outputFilePath, "w", encoding='utf-8') as file:
    for id in sorted_data:
        # This is the data produced by a person.

        suspensionNote = "SUSPENDED / Instance#X / LOST ["
        itemsLost = ""
        itemBarcodes = ""
        # This assigns itemsLost and itemBarcodes as a list on a string for the text file.
        for item in data[id]["Items"]:
            if data[id]["Items"][item]["ProcessStatus"] == "LOST":
                        itemsLost = itemsLost + "\'" + data[id]["Items"][item]["Title"] + "\',"
                        itemBarcodes = itemBarcodes + str(data[id]["Items"][item]["Barcode"]) + ","

        if itemsLost == "":
            # This person does not match the requirements to be suspended, their items are overdue but not lost.
            continue
        else:
            # These cut off the last comma from the lists.
            itemsLost = itemsLost[:len(itemsLost)-1]
            itemBarcodes = itemBarcodes[:len(itemBarcodes)-1]

            suspensionNote = suspensionNote + itemsLost + "]-unresolved- " + str(currentDate.month) + "/" + str(currentDate.day) + "/" + str(currentDate.year) + "-" + initials
            # Legal letter.
            if data[id]["DaysOverdue"] > legalLetterRequirement and breakLine == False:
                breakLine = True
                file.write("\n\n\n\nThe following are people who are eligible to receive a legal letter. Fees are not taken into an account int this list. Requirement: " + str(legalLetterRequirement) + " days overdue.\n\n\n\n")
                 
            ''' Format
            UserId:
            Name: Last, First
            X Days Overdue.
            Suspension note. (You can triple click this line in the text file to easily copy it.)
            Item barcodes in order of the list on the suspension note.
            '''
            file.write("UserId: " + str(id) + "\nName: " + str(data[id]["Name"]) + "\n" + str(data[id]["DaysOverdue"]) + " days overdue.\n" + suspensionNote + "\nThese are the item barcodes in order: " + itemBarcodes + "\n")

# This opens up the exported text file.
try:
    os.system(outputFilePath)
except Exception as e:
    # This has never happened during testing but is a precaution just in case something goes wrong.
    print("An error occured while opening the notepad, but can still be found under the Documents folder.")

# Extends the program's lifetime to view the window just in case it's needed.
print("This window will close in 30 seconds.")
time.sleep(30)