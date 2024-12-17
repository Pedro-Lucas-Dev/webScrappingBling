from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog
import time
import sys

def get_form_data():
    def select_destination():
        global destination_path
        destination_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx")],
            title="Selecione o local para salvar o arquivo"
        )
        label_destination.config(text=destination_path if destination_path else "Nenhum destino selecionado")

    def submit():
        global username, password, selected_store, selected_filter
        username = entry_username.get()
        password = entry_password.get()
        selected_store = store_mapping[filter_store.get()]
        selected_filter = filter_mapping[filter_combobox.get()]
        if not destination_path:
            label_warning.config(text="Por favor, selecione um destino para salvar o arquivo.")
        else:
            root.destroy()

    def on_close():
        root.destroy()
        sys.exit(0)

    root = tk.Tk()
    root.title("Login e Filtro")

    window_width = 500
    window_height = 400
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    position_top = int(screen_height / 2 - window_height / 2)
    position_right = int(screen_width / 2 - window_width / 2)
    root.geometry(f"{window_width}x{window_height}+{position_right}+{position_top}")

    root.protocol("WM_DELETE_WINDOW", on_close)

    tk.Label(root, text="Email:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
    entry_username = tk.Entry(root, width=35)
    entry_username.grid(row=0, column=1, padx=10, pady=10)

    tk.Label(root, text="Senha:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
    entry_password = tk.Entry(root, width=35, show="*")
    entry_password.grid(row=1, column=1, padx=10, pady=10)

    tk.Label(root, text="Selecione uma Loja:").grid(row=2, column=0, padx=10, pady=10, sticky="w")

    tk.Label(root, text="Filtro de Situações:").grid(row=3, column=0, padx=10, pady=10, sticky="w")

    store_mapping = {
        "Mercado Livre Coleta": "204328455",
        "Todas as lojas": "todas",
        "Nenhuma": "0",
        "Amazon 1": "203706129",
        "Amazon FBA Onsite": "204201543",
        "Americanas": "203373883",
        "Faturador Full": "203548425",
        "Físico": "203424099",
        "Loja Integrada": "204753253",
        "Magalu 1": "203614193",
        "Mercado Livre Full": "204338346",
        "Mercado Shop": "204050358",
        "Shopee": "203842216"
    }

    filter_mapping = {
        "Em aberto": "6",
        "Atendido": "9",
        "Cancelado": "12",
        "Em andamento": "15",
        "Venda Agenciada": "18",
        "Em digitação": "21",
        "Verificado": "24",
        "Checkout parcial": "126724",
        "Aguardando Pagamento": "83869",
        "Finalizado": "447110"
    }

    filter_store = ttk.Combobox(root, values=list(store_mapping.keys()), state="readonly", width=32)
    filter_store.grid(row=2, column=1, padx=10, pady=10)
    filter_store.current(0)  

    filter_combobox = ttk.Combobox(root, values=list(filter_mapping.keys()), state="readonly", width=32)
    filter_combobox.grid(row=3, column=1, padx=10, pady=10)
    filter_combobox.current(1)  

    tk.Label(root, text="Destino do Arquivo:").grid(row=4, column=0, padx=10, pady=10, sticky="w")
    label_destination = tk.Label(root, text="Nenhum destino selecionado", width=35, anchor="w")
    label_destination.grid(row=4, column=1, padx=10, pady=10)

    select_button = tk.Button(root, text="Selecionar Destino", command=select_destination, width=20)
    select_button.grid(row=5, column=0, columnspan=2, pady=5)

    label_warning = tk.Label(root, text="", fg="red", wraplength=400)
    label_warning.grid(row=6, column=0, columnspan=2, pady=5)

    submit_button = tk.Button(root, text="Enviar", command=submit, width=15)
    submit_button.grid(row=7, column=0, columnspan=2, pady=10)

    root.mainloop()

get_form_data()

url = "https://www.bling.com.br/login"
newUrl = 'https://www.bling.com.br/vendas.php#list'

driver = webdriver.Chrome()
driver.maximize_window()
driver.get(url)

time.sleep(1)

input_username = driver.find_element(By.ID, "username")
input_username.send_keys(username)
input_password = driver.find_element(By.XPATH, '//input[@type="password"]')
input_password.send_keys(password)

time.sleep(1)
button_login = driver.find_element(By.CLASS_NAME, "login-button-submit")
button_login.click()

time.sleep(3)

driver.get(newUrl)

clear_link = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//a[@class="clr action-link"]'))
)
clear_link.click()

time.sleep(2)

dropdown_store = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "lojasVinculadas-container"))
)
dropdown_store.click()

option_store = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, f"//li[@data-id='{selected_store}']"))
)
option_store.click()

time.sleep(2)

open_filter_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//button[@id="open-filter"]'))
)
open_filter_button.click()

time.sleep(2)

situations_dropdown = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "/html/body/div[8]/div[7]/div[1]/div/form/div[2]/div"))
)
situations_dropdown.click()

time.sleep(2)

situation_filter = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, f"//select[@id='filtro-situacoes']/option[@value='{selected_filter}']"))
)
situation_filter.click()

time.sleep(2)

filter_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//div[@id='filter-button-area']//button[@class='bling-button' and text()='Filtrar']"))
)

driver.execute_script("arguments[0].scrollIntoView();", filter_button)

time.sleep(1)

filter_button.click()

time.sleep(5)

page_source = driver.page_source

soup = BeautifulSoup(page_source, "html.parser")
orders = soup.find_all("span", class_="marcadores")

results = []
for order in orders:
    try:
        order_number_tag = order.find("p", class_="hidden-xs")
        order_number = order_number_tag.text.strip() if order_number_tag else "N/A"
        
        synced_data_tag = order.find("span", title="Dados da nota sincronizados com a loja virtual")
        synced_text = synced_data_tag['title'] if synced_data_tag else ""
        
        results.append({"Pedido": order_number, "Dados Sincronizados": synced_text})
    except Exception as e:
        print(f"Erro ao processar um elemento: {e}")


df = pd.DataFrame(results)
df.to_excel(destination_path, index=False)

driver.quit()
