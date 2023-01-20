import requests
from deep_translator import GoogleTranslator
from bs4 import BeautifulSoup
import pandas as pd
import sys
from datetime import datetime
import time
import re
import openpyxl

dateTimeObj = datetime.now()
filename = f'Finanzen_ref_data_{dateTimeObj.strftime("%d%m%Y-%H%M")}.xlsx'
Finanzendata = []

headers = {
    "User-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36"}
url = 'https://www.finanzen.net/anleihen/{}'
file = open("./FinanzenPortfolio.txt")
lines = file.read().splitlines()
file.close()
#Define variables


for Identifier in lines:
    print(Identifier)

    content = requests.get(url.format(Identifier), headers=headers, stream=True)
    soup = BeautifulSoup(content.text, 'html.parser')
    tbody_data = soup.find_all('tbody')

    if len(tbody_data) <= 1:
        data_status = 'No'
    else:
        data_status = 'Yes'
    # Define variables
    result = {
        'isin': 'n/a',
        'security available': 'n/a',
        'name': 'n/a',
        'wkn': 'n/a',
        'coupon in %': 'n/a',
        'first coupon date': 'n/a',
        'last coupon date': 'n/a',
        'pay coupon': 'n/a',
        'next interest payment date': 'n/a',
        'interest date period': 'n/a',
        'interest dates per year': 'n/a',
        'interest run off': 'n/a',
        'issuer': 'n/a',
        'issue volume': 'n/a',
        'issue currency': 'n/a',
        'issue date': 'n/a',
        'due date': 'n/a',
        'bond type': 'n/a',
        'categorization': 'n/a',
        'issuer group': 'n/a',
        'country': 'n/a',
        'denomination': 'n/a',
        'denomination art': 'n/a',
        'subordinate': 'n/a'
    }

    for elem in tbody_data:
        for rows in elem.find_all("tr"):
            # print ("rows: " + rows.text)
            col = rows.find_all("td")
            data = [ele.text.strip() for ele in col]
            #print(data)
            if (len(data) == 2):
                key = ''
                value = ''
                try:
                  key = GoogleTranslator(source='german', target='en').translate(text=data[0])
                  if key.lower() == 'Surname'.lower():
                      key = 'Name'
                  elif key.lower() == 'issue volume*':
                      key= 'issue volume'
                except:
                    key = data[0]
                try:
                    if value == '-':
                        value = 'no data'
                    else:
                        value = GoogleTranslator(source='german', target='en').translate(text=data[1])
                except:
                    value = data[1]
                result[key.lower()] = value


    result['Security Available'.lower()] = data_status


    def convert_data_format(result, date_cols, date_formats, desired_formats):
        date_dict = {}
        for date_col, date_format, desired_format in zip(date_cols, date_formats, desired_formats):

            try:
                date = datetime.strptime(result[date_col.lower()], date_format).strftime(desired_format)
                date_dict[date_col] = date
            except ValueError as e:
                date_dict[date_col] = result[date_col.lower()]
            except TypeError as e:
                date_dict[date_col] = "Not available"


        return date_dict

    date_cols = ['First Coupon Date', 'Last Coupon Date', 'Next Interest Payment Date','Issue date', 'Due date','Interest Run off']
    date_formats = ['%m/%d/%Y', '%m/%d/%Y', '%m/%d/%Y', '%m/%d/%Y', '%m/%d/%Y', '%m/%d/%Y']
    desired_formats = ['%d.%m.%Y', '%d.%m.%Y', '%d.%m.%Y', '%d.%m.%Y', '%d.%m.%Y','%d.%m.%Y']

    date_dict = convert_data_format(result, date_cols, date_formats, desired_formats)


    tempdata = {'ISIN': Identifier,
                 'Security Available': result['Security Available'.lower()],
                 'Name': result['Name'.lower()],
                 'WKN': result['WKN'.lower()],
                 'Coupon in %': result['Coupon in %'.lower()],
                'First Coupon Date': date_dict['First Coupon Date'],
                 'Last Coupon Date': date_dict['Last Coupon Date'],
                 'Pay Coupon': result['Pay Coupon'.lower()],
                 'Next Interest Payment Date': date_dict['Next Interest Payment Date'],
                 'Interest Date Period': result['Interest Date Period'.lower()],
                 'Interest Dates per year': result['Interest Dates per year'.lower()],
                 'Interest Run off': date_dict['Interest Run off'],
                 'Issuer': result['Issuer'.lower()],
                 'Issue Volume': result['Issue Volume'.lower()],
                 'Issue Currency': result['Issue Currency'.lower()],
                 'Issue date': date_dict['Issue date'],
                 'Due Date': date_dict['Due date'],
                 'Bond Type': result['Bond Type'.lower()],
                 'Categorization': result['Categorization'.lower()],
                 'Issuer Group': result['Issuer Group'.lower()],
                 'Country': result['Country'.lower()],
                 'Denomination': result['Denomination'.lower()],
                 'Denomination Art': result['Denomination Art'.lower()],
                 'Subordinate': result['Subordinate'.lower()],
                 }
    Finanzendata.append(tempdata)
    df = pd.DataFrame.from_dict(Finanzendata)
    df.to_excel(filename, index=False)















