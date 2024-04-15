import os
import time
import datetime
import pandas as pd
import tkinter as tk

from io import StringIO

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains

from dotenv import (
    load_dotenv,
    find_dotenv
)


load_dotenv(find_dotenv())


URL = os.environ.get('URL')

#== Selenium ==#
driver = webdriver.Chrome()


def open_url(url):
    url = os.environ.get('URL')
    driver.get(url)
    time.sleep(10)
    return None


def login(username, password):
    username_input = driver.find_element(By.ID, 'user')
    password_input = driver.find_element(By.ID, 'passw')
    
    username_input.send_keys(username)
    password_input.send_keys(password)
    login_button = driver.find_element(By.ID, 'submit')
    login_button.click()
    time.sleep(12)
    return None


def get_target_url():
    driver.find_element(By.ID, 'hb_mi_apps').click()
    time.sleep(1)
    btn = driver.find_elements(By.CLASS_NAME, 'applications-grid-item')[1]
    btn.click()
    time.sleep(10)
    return None


def data_add_table():
    new_window_handle = [handle for handle in driver.window_handles if handle != driver.current_window_handle][0]
    driver.switch_to.window(new_window_handle)

    actions = ActionChains(driver)
    all_datas = driver.find_elements(By.CLASS_NAME, 'item')
    for i in all_datas:
        actions.double_click(i).perform()

    return None


def table_to_excel():
    table = driver.find_element(By.ID, 'all-stat')
    content = table.get_attribute('outerHTML')
    string_data = StringIO(content)

    df = pd.read_html(string_data)[0]
    df = df[[
        'Unit', 'Rate', 'Duration', 'Mileage, km'
    ]]

    dt = datetime.datetime.today()

    excel_path = f'idrivesafe-data-{dt}.xlsx'
    df.to_excel(excel_path, index=False)

    return None


#== Tkinter ==#
def activate_button():
    button1.config(state=tk.NORMAL)


def enable_button2():
    button2.config(state=tk.NORMAL)


root = tk.Tk()
root.title("Export iDriveSafe")
root.geometry("750x500")


label = tk.Label(root,
                 text="""
                 Zəhmət olmasa bütün məlumatların tam yüklənməsini gözləyin!
                 Bütün məlumatların yüklənməsi tamamlanıbsa, `Məlumatlar yükləndi` klik edin!
                 `Məlumatlar yükləndi` adlı düymə 30 saniyə sonra aktiv olacaq!
                 `Export` adlı düyməni aktiv etmək üçün `Məlumatlar yükləndi`adlı düyməyə klik etmək lazımdır!
                 """)
label.pack(pady=15)

button1 = tk.Button(root, text="Məlumatlar yükləndi", state=tk.DISABLED, command=enable_button2)
button1.pack(pady=20)

button2 = tk.Button(root, text="Export", state=tk.DISABLED, command=table_to_excel)
button2.pack(pady=20)

powered_by = tk.Label(root, text="Powered By Adil", bg="lightgray", fg="black")
powered_by.pack(side=tk.BOTTOM, fill=tk.X)


#== Call Functions ==#
open_url(URL)

login(os.environ.get('TRACKING_USERNAME'), os.environ.get('TRACKING_PASSWORD'))

get_target_url()

data_add_table()

root.after(30000, activate_button)

root.mainloop()

driver.quit()
