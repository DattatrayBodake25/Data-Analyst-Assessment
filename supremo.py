# %%
import concurrent.futures
import csv
import json
import os
import queue
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

import pandas as pd
import pg8000
import psycopg2
from bs4 import BeautifulSoup
from psycopg2.extras import RealDictCursor
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

# Set up Chrome webdriver using ChromeDriverManager
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Main URL to scrape
main_url = 'https://supremo.nic.in/KnowYourOfficerIAS.aspx'
driver.get(main_url)


# Find and select an option (for demonstration purposes)
driver.find_element(By.CLASS_NAME, 'chosen-choices').click()
ul = driver.find_element(By.CLASS_NAME, 'chosen-results')
li = ul.find_elements(By.TAG_NAME, 'li')

for i in range(0, len(li)):
    print(i)
    driver.find_element(By.CLASS_NAME, 'chosen-choices').click()
    time.sleep(0.5)
    li = ul.find_elements(By.TAG_NAME, 'li')
    li[i].click()
    break

# Click the submit button
button = driver.find_element(By.CSS_SELECTOR, '#btnSubmit')
button.click()

# Extract all <a> tags
tr_tags = driver.find_elements(By.TAG_NAME, 'tr')[1:]
a_tags = [tr.find_element(By.TAG_NAME, 'a') for tr in tr_tags]

# Set an implicit wait for the driver
driver.implicitly_wait(10)

# Create a directory to save the tables if it doesn't exist
if not os.path.exists('tables'):
    os.makedirs('tables')


def extract_table(a):
    try:
        a.click()
        driver.switch_to.window(driver.window_handles[1])
        table = driver.find_element(By.TAG_NAME, "table")
        table_html = table.get_attribute('outerHTML')

        # Save the table as an HTML file
        table_index = len(os.listdir('tables')) + 1
        with open(f'tables/table_{table_index}.html', 'w', encoding='utf-8') as f:
            f.write(table_html)
    finally:
        driver.close()
        driver.switch_to.window(driver.window_handles[0])


def save_tables_to_excel():
    tables_folder = 'tables'
    excel_folder = 'excel_tables'

    # Create a directory to save Excel tables if it doesn't exist
    if not os.path.exists(excel_folder):
        os.makedirs(excel_folder)

    # Get a list of HTML files in the tables folder
    html_files = [f for f in os.listdir(tables_folder) if f.endswith('.html')]

    for file in html_files:
        html_path = os.path.join(tables_folder, file)
        with open(html_path, 'r', encoding='utf-8') as html_file:
            table_html = html_file.read()
            soup = BeautifulSoup(table_html, 'html.parser')
            table = soup.find('table')

            # Convert the HTML table to a Pandas DataFrame
            df = pd.read_html(str(table))[0]

            # Save the DataFrame to Excel format
            excel_path = os.path.join(excel_folder, f"{os.path.splitext(file)[0]}.xlsx")
            df.to_excel(excel_path, index=False)
            print(f"Saved {excel_path}")



# Extract tables and save them as HTML
for a in a_tags:
    extract_table(a)

# Save extracted tables to Excel
save_tables_to_excel()

# Quit the driver
driver.quit()