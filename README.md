# CSSkins

A Program to automatically keep Track of your CounterStrike Investments

The Program uses the Steam Community Market Page to track the Value of the Item, and uses the user input Buy Price to calculate the Profit per Item and total Profit.


##Usage
Download the update.exe
Now, create a Excel File Containing your Investments, Amounts and the Buy Price. Make sure the Item Name is the exact Item Name, otherwise it wont work!

Then, create a Excel File Called Investments.xlsx and put it in the same Folder as the Python Script.
Use the Example File or create your own Excel File. It should look like this:

![image](https://user-images.githubusercontent.com/91871891/229320140-3243f65e-8bda-485e-94af-a21a0ee247d3.png)

Run the Exe. If everything was successfull, it should look like this: 

![image](https://user-images.githubusercontent.com/91871891/229320210-ceed2509-c01e-4df2-b6d3-d82cc391f303.png)


## Building it yourself
First, make sure you have Python 3 and pip installed.
Then install the required packages

Open your cmd and execute the following commands:
pip install requests beautifulsoup4 openpyxl forex-python

Now do the Steps required for the Excel File Above
Then, run the update.py file
