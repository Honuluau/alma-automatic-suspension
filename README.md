# Alma automatic suspension

This is an easy to use python script that automatically lists and formats a suspension note of a student with overdue items. The script does not suspend people automatically, human supervision is required. The user must copy and paste the formatted suspension note to the `Add Note` block in Alma.

## How to use.
> [!CAUTION]
> If you have already done this, you do not need to do it again.
> 
> You must have a version of python installed on your local user to run this file. You can find python here: [https://www.python.org/downloads/](https://www.python.org/downloads/)
> 
> This script has 1 dependancy, OpenPyXL. Please run the following command in your "Command Prompt" after installing python:
> ```
> py -m pip install openpyxl
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
1. Once prompted, input your initials into the window.
2. Confirm your initials.
3. Please drag your excel sheet from your downloads folder into the terminal as shown below:  
> <img src="/gifs/drag-and-drop.gif" width="480" height="270"/>  
4. Ensure you have the terminal selected and press enter.

### Using the text file.
The format of each entry goes as follows:  
```
UserId:  
Name: Last, First  
X Days Overdue.  
Suspension note.
Item barcodes in order of the list on the suspension note.   
```
> [!TIP]
> You can triple click the Suspension Note in the text file to easily copy it.  
