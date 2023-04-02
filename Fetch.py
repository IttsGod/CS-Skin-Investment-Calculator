import os
import requests
import openpyxl
import urllib.parse
import re
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, numbers

#Global Variable Declaration
global language
global file_name
global currency
global currency_code
global currency_token
with open('settings.txt', 'r') as f:
    for line in f:
        if line.startswith('language'):
            language = line.split('=')[1].strip()
        elif line.startswith('file_name'):
            file_name = line.split('=')[1].strip()
        elif line.startswith('currency'):
            currency = line.split('=')[1].strip()

if currency.casefold() == "usd":
    currency_code = 1
    currency_token = "$"
elif currency.casefold() == "gbp":
    currency_code = 2
    currency_token = "£"
elif currency.casefold() == "eur":
    currency_code = 3
    currency_token = "€"

def get_skin_price(market_hash_name, session):
    price_overview_url = f"https://steamcommunity.com/market/priceoverview/?appid=730&currency={currency_code}&market_hash_name={market_hash_name}"
    response = session.get(price_overview_url)
    data = response.json()

    if data['success']:
        lowest_price = data['lowest_price']
        cleaned_price = re.sub(r"[^0-9.,]", "", lowest_price).replace(",", ".")
        price = float(cleaned_price)
        return price
    return None

def get_market_hash_name(skin_name, session):
    search_url = f"https://steamcommunity.com/market/search?appid=730&q={skin_name}&l={language}"
    response = session.get(search_url)
    soup = BeautifulSoup(response.text, 'html.parser')

    search_results_row = soup.find('div', class_='market_listing_row market_recent_listing_row market_listing_searchresult')
    if search_results_row:
        market_hash_name = search_results_row['data-hash-name']
        return market_hash_name
    return None

# Load the existing Excel file
excel_file_path = os.path.join(os.getcwd(), file_name)
workbook = load_workbook(excel_file_path)
worksheet = workbook.active

# Create a session
session = requests.Session()

# Iterate through the rows, starting from the second row (skipping the header)
total_profit = 0
for row in range(2, worksheet.max_row  + 1):
    skin_name = worksheet.cell(row=row, column=1).value
    if skin_name is not None:
        skin_name = urllib.parse.quote(skin_name)
        market_hash_name = get_market_hash_name(skin_name, session)
        if market_hash_name is not None:
            market_hash_name = urllib.parse.quote(market_hash_name)
            price = get_skin_price(market_hash_name, session)
            # rest of the code
            if price is not None:
                cell = worksheet.cell(row=row, column=4, value=price)
                cell.number_format = f'#,##0.00\ "{currency_token}";[Red]\-#,##0.00\ "{currency_token}"'
                purchase_price = worksheet.cell(row=row, column=3).value
                profit = price - purchase_price
                profit_cell = worksheet.cell(row=row, column=5, value=profit)
                profit_cell.number_format = f'#,##0.00\ "{currency_token}";[Red]\-#,##0.00\ "{currency_token}"'
                amount = worksheet.cell(row=row, column=2).value
                total_profit_per_item = profit * amount
                total_profit += total_profit_per_item
                total_profit_cell = worksheet.cell(row=row, column=6, value=total_profit_per_item)
                total_profit_cell.number_format = f'#,##0.00\ "{currency_token}";[Red]\-#,##0.00\ "{currency_token}"'


            else:
                worksheet.cell(row=row, column=4, value="Not Found")
                print("Couldnt find Price for Item: " + skin_name + ". Please check your Spelling, and if this happens often, try again in 2 Minutes")

all_total_profit_cell = worksheet.cell(row=2, column=7, value=total_profit)
all_total_profit_cell.number_format = f'#,##0.00\ "{currency_token}";[Red]\-#,##0.00\ "{currency_token}"'

# Save the updated Excel file
workbook.save(excel_file_path)