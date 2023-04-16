import requests
from bs4 import BeautifulSoup
import pyodbc
from datetime import datetime
from selenium import webdriver
import schedule
import time

def scrape_data():
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=E:\data\python\KirimSaja\kirim.mdb;')
    cursor = conn.cursor()

     # Cek apakah ada data pada tanggal ini
    query = "SELECT COUNT(*) FROM DataSahamx WHERE Tanggal=?"
    count = cursor.execute(query, date.today().strftime("%Y-%m-%d")).fetchone()[0]

    if count == 0:
        url = 'http://www.infovesta.com/index2/shlq45'
        driver = webdriver.Chrome()
        driver.get(url)
        response = driver.page_source
        with open('data.html', 'w', encoding='utf-8') as f:
            f.write(response)
        driver.quit()
        
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
        print("Data berhasil disimpan pada", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    else:
        print("Data pada tanggal ini sudah tersimpan di database.")

# Fungsi untuk menjalankan program hanya pada hari Senin, Selasa, Rabu, Kamis, dan Jumat pada pukul 20:00 WIB
def run_scraping_job():
    day = datetime.today().weekday()
    hour = datetime.now().strftime("%H")
    if day in [0,1,2,3,4] and hour == '20':
        scrape_data()

# Menjadwalkan program untuk dijalankan setiap 1 menit sekali
schedule.every(1).minutes.do(run_scraping_job)

while True:
    schedule.run_pending()
    time.sleep(1)