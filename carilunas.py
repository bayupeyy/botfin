from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import pandas as pd
import time

# Lokasi WebDriver dan file
CHROMEDRIVER_PATH = r"D:\carilunas\chromedriver.exe"
file_path = r"D:\carilunas\Dummy.xls"
output_file_path_belum_lunas = r"D:\carilunas\Dummy_belum_lunas.xlsx"

# Baca file Excel
data = pd.read_excel(file_path)

# Inisialisasi Service untuk ChromeDriver dengan opsi headless
chrome_options = Options()
chrome_options.add_argument("--headless")  # Menyembunyikan UI browser
chrome_options.add_argument("--disable-gpu")  # Menghindari masalah pada sistem tanpa GPU
service = Service(CHROMEDRIVER_PATH)
driver = webdriver.Chrome(service=service, options=chrome_options)

# List untuk menyimpan data yang belum lunas
data_belum_lunas = []

# Total jumlah data
total_data = len(data)

try:
    # Iterasi setiap baris data
    for index, row in data.iterrows():
        customer_no = row['customerNo']  # Sesuaikan nama kolom jika berbeda

        # Buka website
        driver.get("https://live.finpay.id/widgetpg/001001")
        
        # Tunggu hingga input field muncul
        try:
            input_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder="Nomor Pelanggan"]'))
            )
        except TimeoutException:
            print(f"Timeout: Tidak dapat menemukan input untuk customerNo {customer_no}")
            continue

        # Masukkan customerNo ke input field
        input_field.clear()  # Pastikan input field kosong sebelum menambahkan nomor
        input_field.send_keys(str(customer_no))

        # Klik tombol "Lanjutkan"
        button = driver.find_element(By.XPATH, "//button[contains(., 'Lanjutkan')]")
        button.click()

        # Tunggu hasil
        try:
            # Tunggu elemen terima kasih
            lunas_message = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//p[contains(text(), 'Terima kasih, pembayaran Anda')]"))
            )
            print(f"Tagihan untuk customerNo {customer_no} lunas. Menghapus dari file Excel.")
            data.drop(index, inplace=True)  # Hapus baris jika dianggap lunas
        except TimeoutException:
            print(f"Tagihan untuk customerNo {customer_no} belum lunas.")
            # Menambahkan data yang belum lunas ke list
            data_belum_lunas.append(row)

        # Menampilkan progres pengecekan di terminal
        processed_data = index + 1
        progress = (processed_data / total_data) * 100
        print(f"Progres: {progress:.2f}% | Sudah diperiksa: {processed_data}/{total_data} | Belum diperiksa: {total_data - processed_data}")

# Perintah ketika program dihentikan
except KeyboardInterrupt:
    print("\nProgram dihentikan oleh pengguna. Menyimpan hasil sementara...")

finally:
    # Tutup browser
    driver.quit()
    
# Tambahkan angka 0 di depan kolom telp1
if 'telp1' in data.columns:  # Pastikan kolom telp1 ada
    data['telp1'] = data['telp1'].apply(lambda x: f"0{x}" if pd.notnull(x) else x)

# Simpan data yang belum lunas ke file baru
if data_belum_lunas:
    df_belum_lunas = pd.DataFrame(data_belum_lunas)
    # Tambahkan angka 0 di depan kolom telp1 untuk data yang belum lunas
    if 'telp1' in df_belum_lunas.columns:
        df_belum_lunas['telp1'] = df_belum_lunas['telp1'].apply(lambda x: f"0{x}" if pd.notnull(x) else x)
    df_belum_lunas.to_excel(output_file_path_belum_lunas, index=False)

print(f"File hasil telah disimpan ke {output_file_path_lunas}")
if data_belum_lunas:
    print(f"File hasil yang belum lunas telah disimpan ke {output_file_path_belum_lunas}")
