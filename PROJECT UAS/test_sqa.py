from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time

# Setup driver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.maximize_window()

URL = "http://127.0.0.1:5500/index.html"

def ambil_screenshot(nama_file):
    driver.save_screenshot(f"{nama_file}.png")
    print(f"Screenshot disimpan: {nama_file}.png")

def reset_form():
    driver.get(URL)
    time.sleep(1)

# =============================================
# TEST 1 — Entri Data Karakter (Agus)
# =============================================
reset_form()
input_field = driver.find_element(By.ID, "namaInput")
input_field.send_keys("Agus")
driver.find_element(By.TAG_NAME, "button").click()
time.sleep(1)
ambil_screenshot("test1_karakter")
print("Test 1 - Karakter:", driver.find_element(By.ID, "pesan").text)

# =============================================
# TEST 2 — Entri Data Number (123)
# =============================================
reset_form()
input_field = driver.find_element(By.ID, "namaInput")
input_field.send_keys("123")
driver.find_element(By.TAG_NAME, "button").click()
time.sleep(1)
ambil_screenshot("test2_number")
print("Test 2 - Number:", driver.find_element(By.ID, "pesan").text)

# =============================================
# TEST 3 — Entri Data Karakter+Number (Agus123)
# =============================================
reset_form()
input_field = driver.find_element(By.ID, "namaInput")
input_field.send_keys("Agus123")
driver.find_element(By.TAG_NAME, "button").click()
time.sleep(1)
ambil_screenshot("test3_kombinasi")
print("Test 3 - Kombinasi:", driver.find_element(By.ID, "pesan").text)

# =============================================
# TEST 4 — Input Kosong (Negative Testing)
# =============================================
reset_form()
driver.find_element(By.TAG_NAME, "button").click()
time.sleep(1)
ambil_screenshot("test4_kosong")
print("Test 4 - Kosong:", driver.find_element(By.ID, "pesan").text)

# =============================================
# Selesai
# =============================================
print("\n✅ Semua pengujian selesai! Cek file PNG di folder yang sama.")
time.sleep(2)
driver.quit()