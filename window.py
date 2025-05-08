from tkinter import *
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
import time
import urllib


def btn_clicked():
    funcionario = entry1.get()
    msg = entry0.get("1.0", "end-1c")

    tabela = pd.read_excel('Cobrancas.xlsx')

    nav = webdriver.Chrome()
    nav.get('https://web.whatsapp.com/')

    # fazendo login no whatsapp web
    while len(nav.find_elements_by_id('pane-side')) == 0:
        time.sleep(10)

    # enviando mensagem para o devedor
    for i, situacao in enumerate(tabela['Situação Pagamento']):
        if situacao in msg:
            nome = tabela.loc[i, 'Nome']
            telefone = tabela.loc[i, 'Telefone']   
            valor = tabela.loc[i, 'Valor Devido']
            mensagem = f"Olá {nome}, tudo bem? Você possui um valor de R$ {valor} em aberto. Por favor, entre em contato para regularizar sua situação."
            texto = urllib.parse.quote(mensagem)
            nav.get(f'https://api.whatsapp.com/send?phone={telefone}&text={texto}')

            while len(nav.find_elements_by_id('pane-side')) == 0:
                time.sleep(10)
            time.sleep(5)
            nav.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[1]/div[2]/div[1]/p').send_keys(Keys.ENTER)
            time.sleep(5)
            tabela.loc[i, 'Situação Pagamento'] = 'Mensagem Enviada'
    nav.quit()
    tabela.to_excel('Cobrancas2.xlsx', index=False)  


window = Tk()

window.geometry("616x357")
window.configure(bg = "#ffffff")
canvas = Canvas(
    window,
    bg = "#ffffff",
    height = 357,
    width = 616,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge")
canvas.place(x = 0, y = 0)

background_img = PhotoImage(file = f"background.png")
background = canvas.create_image(
    311.0, 178.5,
    image=background_img)

img0 = PhotoImage(file = f"img0.png")
b0 = Button(
    image = img0,
    borderwidth = 0,
    highlightthickness = 0,
    command = btn_clicked,
    relief = "flat")

b0.place(
    x = 426, y = 301,
    width = 71,
    height = 32)

entry0_img = PhotoImage(file = f"img_textBox0.png")
entry0_bg = canvas.create_image(
    465.5, 244.5,
    image = entry0_img)

entry0 = Text(
    bd = 0,
    bg = "#ffffff",
    highlightthickness = 0)

entry0.place(
    x = 381.0, y = 203,
    width = 169.0,
    height = 81)

entry1_img = PhotoImage(file = f"img_textBox1.png")
entry1_bg = canvas.create_image(
    465.5, 153.0,
    image = entry1_img)

entry1 = Entry(
    bd = 0,
    bg = "#ffffff",
    highlightthickness = 0)

entry1.place(
    x = 381.0, y = 137,
    width = 169.0,
    height = 30)

window.resizable(False, False)
window.mainloop()
