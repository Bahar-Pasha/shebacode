

import requests
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import csv
import pandas as pd

# Define the URL of the web page
url = 'https://cardinfo.ir/info/sheba_decoder'

# name of the Excel file
filename = 'Sheba.xlsx'

# name of the column to extract
column_name = 'Sheba Codes'
# read the data from the Excel file
df = pd.read_excel('Sheba.xlsx')

# read the data from the Excel file
df = pd.read_excel('Sheba.xlsx')

# extract the column data
column_data = df['Sheba Codes']


items = []

# Create a new instance of the Chrome driver
driver = webdriver.Chrome()

# Load the web page
driver.get(url)

# Wait for the page to finish loading
wait = WebDriverWait(driver, 10)
wait.until(EC.presence_of_element_located((By.ID, "sheba_decode_input")))

for input_string in column_data:
    # Find the input element and set its value to the input string
    input_element = driver.find_element(By.ID, 'sheba_decode_input')
    input_element.clear()  # clear the input field
    input_element.send_keys(input_string)
    input_element.send_keys(Keys.RETURN)

    # Wait for the result to appear
    wait.until(EC.visibility_of_element_located((By.ID, "sheba_decode_result")))

    # Find the span element within the div element with id="sheba_decode_result" and extract the text string
    result_div = driver.find_element(By.ID, 'sheba_decode_result')
    #result_span = result_div.find_element(By.TAG_NAME, 'span')
    result_string = result_div.text

    # Store the result in a dictionary
    item = {'Sheba': input_string, 'Account': result_string}
    items.append(item)

# Print the results
for item in items:
    print(item)
# name of the output file

filename = 'Data.csv'


# open the file in write mode
with open(filename, 'w',encoding="utf-8-sig", newline='') as file:
    # create a CSV writer object
    fields = ['Sheba', 'Account']
    writer = csv.DictWriter(file,fieldnames=fields)


    writer.writeheader()
    # write the data to the file
    
    writer.writerows(items)

print('Data exported to', filename)

# Close the browser window
driver.quit()