import requests
from deep_translator import GoogleTranslator
from bs4 import BeautifulSoup
import pandas as pd
import time
import random
from datetime import datetime


# check if required data is available on website
def get_data_status(tbody_data):
    if len(tbody_data) <= 1:
        return 'No'
    else:
        return 'Yes'


# extract benchmark value
def extract_benchmark_value(tbody_data):
    benchmark_value = None
    found_Kupon = False
    for tbody in tbody_data:
        tr_elements = tbody.find_all('tr')
        for idx, row in enumerate(tr_elements):
            cols = row.find_all('td')
            if len(cols) == 2:
                if cols[0].get_text(strip=True) == 'Kupon in %':
                    found_Kupon = True
                elif cols[0].get_text(strip=True) == 'Emittent':
                    found_Kupon = False
                    break
                if found_Kupon and '...Hinweis' in cols[0].get_text(strip=True):
                    benchmark_value = cols[1].get_text().strip()
                    break
        if benchmark_value:
            break
    return benchmark_value


## translate text from German to English

def translation_text(text):
    try:
        datetime.strptime(text, '%d.%m.%Y')
        return text
    except ValueError:
        pass  # Not a date, proceed with translation
    try:
        return GoogleTranslator(source='german', target='en').translate(text=text)
    except:
        return text

# date format conversion

# def convert_data_format(result, date_cols, date_formats, desired_formats):
#     date_dict = {}
#     for date_col, date_format, desired_format in zip(date_cols, date_formats, desired_formats):
#         try:
#             date = datetime.strptime(result[date_col.lower()], date_format).strftime(desired_format)
#             date_dict[date_col] = date
#         except (ValueError , TypeError) :
#             date_dict[date_col] = "Not available"
#     return date_dict

# Main program

start = time.time()

dateTimeObj = datetime.now()
filename = f'Finanzen_ref_data_{dateTimeObj.strftime("%d%m%Y-%H%M")}.xlsx'
Finanzendata = []

headers = {
"User-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36"
}
url = 'https://www.finanzen.net/anleihen/{}'
file = open('./FinanzenPortfolio.txt')
lines = file.read().splitlines()
file.close()

for Identifier in lines:
    print(Identifier)
    #Genearte a random delay
    random_delay = random.uniform(1,5)
    time.sleep(random_delay)
    content = requests.get(url.format(Identifier), headers=headers, stream=True)
    soup = BeautifulSoup(content.text, 'html.parser')
    tbody_data = soup.find_all('tbody')

    data_status = get_data_status(tbody_data)
    benchmark_value = extract_benchmark_value(tbody_data)

    # Define variables
    result = {
        'isin': 'n/a',
        'security available': 'n/a',
        'name': 'n/a',
        'wkn': 'n/a',
        'coupon in %': 'n/a',
        '...a notice': 'n/a',
        'first coupon date': 'n/a',
        'last coupon date': 'n/a',
        'payment coupon': 'n/a',
        'next interest payment date': 'n/a',
        'interest date period': 'n/a',
        'interest dates per year': 'n/a',
        'interest run': 'n/a',
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
        'subordination': 'n/a'
    }

    # Loop through data rows to populate result dictionary
    for elem in tbody_data:
        for rows in elem.find_all("tr"):
            # print ("rows: " + rows.text)
            col = rows.find_all("td")
            data = [ele.text.strip() for ele in col]
            if (len(data) == 2):
                key = translation_text(data[0])
                if key.lower() == 'Surname'.lower():
                    key = 'Name'
                elif key.lower() == 'issue volume*':
                    key = 'issue volume'
                value = translation_text(data[1])
                if value == '-':
                    value = 'no data'
                result[key.lower()] = value

    result['Security Available'.lower()] = data_status

    # date_cols = ['First Coupon Date', 'Last Coupon Date', 'Next Interest Payment Date', 'Issue date', 'Due date',
    #              'Interest Run off']
    # date_formats = ['%m/%d/%Y', '%m/%d/%Y', '%m/%d/%Y', '%m/%d/%Y', '%m/%d/%Y', '%m/%d/%Y']
    # desired_formats = ['%d.%m.%Y', '%d.%m.%Y', '%d.%m.%Y', '%d.%m.%Y', '%d.%m.%Y', '%d.%m.%Y']

    # date_dict = convert_data_format(result, date_cols, date_formats, desired_formats)

    tempdata = {'ISIN': Identifier,
                'Security Available': result['Security Available'.lower()],
                'Name': result['Name'.lower()],
                'WKN': result['WKN'.lower()],
                'Coupon in %': result['Coupon in %'.lower()],
                'Benchmark': benchmark_value,
                'First Coupon Date': result['First Coupon Date'.lower()],
                'Last Coupon Date': result['Last Coupon Date'.lower()],
                'Pay Coupon': result['Payment Coupon'.lower()],
                'Next Interest Payment Date': result['Next Interest Payment Date'.lower()],
                'Interest Date Period': result['Interest Date Period'.lower()],
                'Interest Dates per year': result['Interest Dates per year'.lower()],
                'Interest Run off': result['Interest Run'.lower()],
                'Issuer': result['Issuer'.lower()],
                'Issue Volume': result['Issue Volume'.lower()],
                'Issue Currency': result['Issue Currency'.lower()],
                'Issue date': result['Issue date'.lower()],
                'Due Date': result['Due date'.lower()],
                'Bond Type': result['Bond Type'.lower()],
                'Categorization': result['Categorization'.lower()],
                'Issuer Group': result['Issuer Group'.lower()],
                'Country': result['Country'.lower()],
                'Denomination': result['Denomination'.lower()],
                'Denomination Art': result['Denomination Art'.lower()],
                'Subordinate': result['Subordination'.lower()],
                }
    Finanzendata.append(tempdata)
    # print(Finanzendata)
df = pd.DataFrame.from_dict(Finanzendata)
df.to_excel(filename, index=False)
print('It took', (time.time() - start) / 60, 'minutes.')
