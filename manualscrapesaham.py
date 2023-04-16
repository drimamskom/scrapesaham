import requests
from bs4 import BeautifulSoup
import pyodbc
from datetime import datetime
from selenium import webdriver

url = 'http://www.infovesta.com/index2/shlq45'
driver = webdriver.Chrome()
driver.get(url)
response = driver.page_source
with open('data.html', 'w', encoding='utf-8') as f:
    f.write(response)

driver.quit()
conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=E:\data\python\KirimSaja\kirim.mdb;')
cursor = conn.cursor()
query = "INSERT INTO DataSahamx (Kode, Tanggal, open, high, low, close, 1hr, 1bln, 1thn) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"
with open('data.html', 'r') as f:
    html = f.read()
soup = BeautifulSoup(html, 'html.parser')
tbodies = soup.find_all('tbody', {'role': 'rowgroup'})
for tbody in tbodies:
    for row in tbody.find_all('tr'):
        cols = row.find_all('td')
        cols = [col.text.strip() for col in cols]
        cols.insert(1, datetime.now().strftime("%Y-%m-%d")) # Menambahkan tanggal saat ini ke dalam list
        print(cols)
        cursor.execute(query, cols)
conn.commit()
cursor.close()
conn.close()