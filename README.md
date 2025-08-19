# Alma automatic suspension
Github repository: [https://github.com/Honuluau/alma-automatic-suspension](https://github.com/Honuluau/alma-automatic-suspension)

This is an easy-to-use python script that automatically lists and formats a suspension note of a student with overdue items. The script does not suspend people automatically, human supervision is required. The user must copy and paste the formatted suspension note to the `Add Note` block in Alma.

## How to use.
> [!CAUTION]
> This is a NEW library that is DIFFERENT from the PREVIOUS.
> If you have already done this, you do not need to do it again.
> 
> You must have a version of python installed on your local user to run this file. You can find python here: [https://www.python.org/downloads/](https://www.python.org/downloads/)
> 
> This script has 1 dependency, pandas. Please run the following command in your "Command Prompt" after installing python:
> ```
> py -m pip install pandas
> ```

### Downloading the Alma Item Report
Please go to your Alma Dashboard.
1. Go to the Analytics page which can be found on the left side of the screen.
2. Go to "Out of the Box Analytics/Reports".
3. Look for "Fulfillment - Loans Returns and Overdue Dashboard (Ex Libris)" and select it.
4. Click the blue button "View Full Report"
5. Once loaded, click the top right gear icon.
6. Hover over export to excel and click "Export Entire Dashboard"
7. You can find this file under downloads in the file explorer.

### Using the python script.
Open the python (.py) file that is already in EVE 2.0 or the one downloaded from this github page. 
1. Wait for a file explorer to pop up that asks for a fulfillment report.
2. Select the most recent fulfillment report downloaded and click "Open".
3. Enter your initials when prompted.
4. Confirm your initials.
5. A spreadsheet should appear on Microsoft Excel. 
> [!CAUTION]
> If you do not have access to Excel or this does not occur, navigate to your Documents and find the "Alma-Automatic-Suspensions Logs" folder. The most recent file by timestamp should be the output.

### Using the spreadsheet file.
The format of the spreadsheet is as follows: 
```
Eagle_Id, Last Name First Name, Longest Overdue (Days), Number of Items Lost, Suspension Note
```
The formatting on the Eagle_Id may look strange, but copying it should paste the appropriate data.
Copy the Eagle Id into Alma and search for that user.

Select the matching note under the "E" or "Suspension Note" Column by clicking on "SUSPENDED" and copy it.
On the left column, select "Manage notes" and select "Add Note".
In the Add Note section on Alma, paste this note into it.

> [!CAUTION]
> If "Add Note" is missing, it's because the user does not have the permissions to manage notes.
> 
> If you are unable to copy and paste from Microsoft Excel, it is due to your permissions. Make sure to be signed in and have no problems with your Microsoft Account. You can manage this from the top right circle icon on Excel.