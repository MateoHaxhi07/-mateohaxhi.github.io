import os
import time
import datetime
import glob
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ---------------- CONSTANTS ----------------
LOGIN_URL = 'https://hospitality.devpos.al/login'
REPORTS_URL = 'https://hospitality.devpos.al/user/0/produktet/shitjet'
NIPT = "K31412026L"
USERNAME = "Elona"
PASSWORD = "Sindi2364*"

EXISTING_FILE = r"C:\Users\mhaxh\OneDrive\Desktop\2025\data\sales_till_01_31.xlsx"  
def setup_driver():
    script_dir = os.path.abspath(os.path.dirname(__file__))
    download_folder = os.path.join(script_dir, "data")

    if not os.path.exists(download_folder):
        os.makedirs(download_folder)

    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": download_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=chrome_options)
    return driver

def login_to_website(driver):
    driver.get(LOGIN_URL)
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.NAME, 'nipt'))
    ).send_keys(NIPT)

    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.NAME, 'username'))
    ).send_keys(USERNAME)
    
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//input[@formcontrolname="password"]'))
    ).send_keys(PASSWORD)

    driver.find_element(By.XPATH, "//button[contains(., 'Login')]").click()
    time.sleep(5)  # Wait for any post-login processes

def get_unique_filename(base_path):
    if not os.path.exists(base_path):
        return base_path
    
    base, ext = os.path.splitext(base_path)
    count = 1
    while True:
        new_path = f"{base}_({count}){ext}"
        if not os.path.exists(new_path):
            return new_path
        count += 1

def format_excel_file(file_path):
    df_new = pd.read_excel(file_path)

    # Delete unnecessary columns
    columns_to_delete = [0, 1, 2, 5, 7, 8, 9, 10, 12, 13, 15, 16, 18, 20, 21, 23, 24, 25]
    df_new.drop(df_new.columns[columns_to_delete], axis=1, inplace=True)

    df_new['Data Rregjistrimit'] = pd.to_datetime(df_new['Data Rregjistrimit'], format='%d/%m/%Y', errors='coerce')
    df_new['Koha Rregjistrimit'] = df_new['Koha Rregjistrimit'].astype(str)
    df_new['Datetime'] = pd.to_datetime(df_new['Data Rregjistrimit'].dt.strftime('%Y-%m-%d') + ' ' + df_new['Koha Rregjistrimit'], errors='coerce')

    df_new.drop(['Data Rregjistrimit', 'Koha Rregjistrimit'], axis=1, inplace=True)

    # Rename columns
    new_column_names = {
        df_new.columns[0]: 'Seller',
        df_new.columns[1]: 'Article_Name',
        df_new.columns[2]: 'Category',
        df_new.columns[3]: 'Quantity',
        df_new.columns[4]: 'Article_Price',
        df_new.columns[5]: 'Total_Article_Price'
    }
    df_new.rename(columns=new_column_names, inplace=True)

    df_new = process_data(df_new)  # Process the data (Seller Category, etc.)

    # ✅ **Append to the existing file**
    if os.path.exists(EXISTING_FILE):
        df_existing = pd.read_excel(EXISTING_FILE)

        # Avoid duplicate rows before appending
        df_combined = pd.concat([df_existing, df_new]).drop_duplicates()

        df_combined.to_excel(EXISTING_FILE, index=False)
        print(f"Appended new data to: {EXISTING_FILE}")
    else:
        df_new.to_excel(EXISTING_FILE, index=False)
        print(f"Created new file: {EXISTING_FILE}")

def process_data(df):
    seller_categories = {
        'Enisa': 'Delivery',
        'Dea': 'Delivery',
        'Kristian Llupo': 'Bar',
        'Pranvera Xherahi': 'Bar',
        'Fjorelo Arapi': 'Restaurant',
        'Jonel Demba': 'Restaurant'
    }
    
    df['Seller Category'] = df['Seller'].map(seller_categories)
    df = df[df['Seller'] != 'TOTALI']
    return df

def cleanup_excel_files(download_folder):
    excel_files = [os.path.join(download_folder, f) for f in os.listdir(download_folder) if f.endswith('.xlsx')]
    if not excel_files:
        print("No Excel files found in the directory.")
        return

    newest_file = max(excel_files, key=os.path.getmtime)
    print(f"Keeping the newest file: {newest_file}")

    for file in excel_files:
        if file != newest_file:
            os.remove(file)
            print(f"Removed file: {file}")

def download_excel_report(driver):
    driver.get(REPORTS_URL)
    WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Shkarko raportin')]"))
    ).click()

    download_folder = os.path.join(os.path.abspath(os.path.dirname(__file__)), "data")
    while True:
        matching_files = glob.glob(os.path.join(download_folder, "raport shitjes*.xlsx"))
        if matching_files:
            file_path = matching_files[0]
            if not file_path.endswith(".crdownload"):
                time.sleep(1)
                today_str = datetime.now().strftime("%m_%d")
                base_new_filename = f"sales_till_{today_str}.xlsx"
                new_file_path = os.path.join(download_folder, base_new_filename)
                try:
                    os.rename(file_path, new_file_path)
                    print(f"File renamed to: {new_file_path}")
                    format_excel_file(new_file_path)  # Append data instead of replacing
                    cleanup_excel_files(download_folder)
                    break
                except PermissionError as e:
                    print(f"PermissionError during file rename: {e}")
                    time.sleep(2)
        time.sleep(1)

def main():
    driver = setup_driver()
    try:
        login_to_website(driver)
        download_excel_report(driver)
        print("Current Timestamp:", datetime.datetime.now())
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
