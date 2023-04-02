# CSSkins

A Program to automatically keep track of your CounterStrike Investments

The Program uses the Steam Community Market page to track the Value of the Item, and uses the user input buy price to calculate the profit per item and total profit.
It does not require or send any User / Account Data, and converts the currency 


## Usage

Download the update.exe, and the settings.txt files from the latest release.

Now, create a Excel file containing your investments, Amounts and the buy price. Make sure the item name is the exact item name, otherwise it wont work!

### Settings Syntax

The following Settings can be changed in the settings.txt file

language: The written Language, examples are: english, german, french, etc.

file_name: The name of the file (and the location, if your file is one directory above your script, you can use ..\filename.xlsx for example)

currency: The Currency the Program checks. Currently supported: USD, GBP, EUR

### Excel Setup

The Excel file should be in the same folder as your executable.

Use the Example File or create your own Excel File. It should look like this:

![image](https://user-images.githubusercontent.com/91871891/229320140-3243f65e-8bda-485e-94af-a21a0ee247d3.png)

Run the executable. If everything was successfull, it should look like this: 

![image](https://user-images.githubusercontent.com/91871891/229320210-ceed2509-c01e-4df2-b6d3-d82cc391f303.png)


## Building it yourself
First, make sure you have Python 3 and pip installed.
Then install the required packages

Open your cmd and execute the following command:
`pip install requests beautifulsoup4 openpyxl forex-python`

Now do the Steps required for the Excel file above
Then, run the update.py file
