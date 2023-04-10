# CS Investment Calculator

A Program to automatically keep track of your CounterStrike Investments

First of all, if you have any Questions, please feel free to DM me on Discord about them. My Discord is Arik_#9099

The Program uses the Steam Community Market page to track the Value of the Item, and uses the user input buy price to calculate the profit per item and total profit.
It does not require or send any User / Account Data.


## Usage

Download the update.exe, and the settings.txt files from the latest release.

Now, create a Excel file containing your investments, Amounts and the buy price. Make sure the item name is the exact item name, otherwise it wont work!

### Settings Syntax

The following Settings can be changed in the settings.txt file

language: The language in which the items are listed. Examples are English, German, French, etc. 

file_name: The name of the file (and the location, if your file is one directory above your script, you can use ..\filename.xlsx for example)

currency: The Currency the Program checks. Currently supported: USD, GBP, EUR

update_hours: The Hours, after which the Skins should be updated if the File is run. 

### Excel Setup

The Excel file should be in the same folder as your executable.

Use the Example File or create your own Excel File. It should look like this:

![image](https://user-images.githubusercontent.com/91871891/230908601-e1579dc9-eede-416f-ac55-fc71508ddd98.png)

Run the executable. If everything was successfull, it should look like this: 

![image](https://user-images.githubusercontent.com/91871891/230908561-15038877-1c42-4ab5-9d9c-618e6613683d.png)

You can then use Excel to Calculate your whole Profit, whole Price, or something else in the Columns right to the timestamp
The File will also create a second Spreadsheet, please do not change that.

## Building it yourself
First, make sure you have Python 3 and pip installed.
Then install the required packages

Open your cmd and execute the following command:
`pip install requests openpyxl urllib beautifulsoup4 re`

Now do the Steps required for the Excel file above
Then, run the update.py file
