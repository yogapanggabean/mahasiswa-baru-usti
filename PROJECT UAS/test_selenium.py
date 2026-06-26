from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import os

# buka file HTML lokal kamu
file_path = "file:///" + os.path.abspath("index.html")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.get(file_path)

time.sleep(2)

# =========================
# TEST 1 - KARAKTER
# =========================
input_box = driver.find_element(By.ID, "nama")
input_box.clear()
input_box.send_keys("Agus")
input_box.send_keys(Keys.ENTER)

time.sleep(2)
driver.save_screenshot("test_karakter.png")

# =========================
# TEST 2 - NUMBER
# =========================
input_box = driver.find_element(By.ID, "nama")
input_box.clear()
input_box.send_keys("123")
input_box.send_keys(Keys.ENTER)

time.sleep(2)
driver.save_screenshot("test_number.png")

# =========================
# TEST 3 - KOMBINASI
# =========================
input_box = driver.find_element(By.ID, "nama")
input_box.clear()
input_box.send_keys("Agus123")
input_box.send_keys(Keys.ENTER)

time.sleep(2)
driver.save_screenshot("test_kombinasi.png")

# =========================
# TEST 4 - KOSONG
# =========================
input_box = driver.find_element(By.ID, "nama")
input_box.clear()
input_box.send_keys("")
input_box.send_keys(Keys.ENTER)

time.sleep(2)
driver.save_screenshot("test_kosong.png")

driver.quit()