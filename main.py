import tkinter as tk
from tkinter import ttk
import numpy as np
from tkcalendar import DateEntry
from tkinter.filedialog import askopenfilename
import requests
import pandas as pd
from datetime import datetime


requisicao = requests.get('https://economia.awesomeapi.com.br/json/all')
dictionary_currencies = requisicao.json()

currency_list = list(dictionary_currencies.keys())


def pegar_cotacao():
    moeda = combobox_selectcurrency.get()
    data_cotacao = currency_calendar.get()
    ano = data_cotacao[-4:]
    mes = data_cotacao[3:5]
    dia = data_cotacao[:2]
    link = f'https://economia.awesomeapi.com.br/json/daily/{moeda}/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}'
    requisicao_moeda = requests.get(link)
    cotacao = requisicao_moeda.json()
    valor_moeda = cotacao[0]['bid']
    currency_text['text'] = f"A Cotação da {moeda} no dia {data_cotacao} foi de: R${valor_moeda}"


def selecionar_arquivo():
    caminho_arquivo = askopenfilename(title='Selecione o Arquivo de Moedas')
    var_pathfile.set(caminho_arquivo)
    if caminho_arquivo:
        selected_file['text'] = f'Arquivo Selecionado: {caminho_arquivo}'


def atualizar_cotacoes():
    try:
        df = pd.read_excel(var_pathfile.get())
        moedas = df.iloc[:, 0]
        data_inicial = calendar_start_date.get()
        data_final = calendar_end_date.get()
        ano_inicial = data_inicial[-4:]
        mes_inicial = data_inicial[3:5]
        dia_inicial = data_inicial[:2]

        ano_final = data_final[-4:]
        mes_final = data_final[3:5]
        dia_final = data_final[:2]

        for moeda in moedas:
            link = link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/31?" \
               f"start_date={ano_inicial}{mes_inicial}{dia_inicial}" \
               f"&end_date={ano_final}{mes_final}{dia_final}"

            requisicao_moeda = requests.get(link)
            cotacoes = requisicao_moeda.json()
            for cotacao in cotacoes:
                timestamp = int(cotacao['timestamp'])
                bid = float(cotacao['bid'])
                data = datetime.fromtimestamp(timestamp)
                data = data.strftime('%d/%m/%Y')
                if data not in df:
                    df[data] = np.nan

                df.loc[df.iloc[:, 0] == moeda, data] = bid
        df.to_excel('Moedas.xlsx')
        update_quotation['text'] = 'Arquivo Atualizado com Sucesso'
    except:
        update_quotation['text'] = 'Selecione um arquivo Excel'


window = tk.Tk()
window.title('Ferramenta de Cotação de Moeda')

currency_quotation = tk.Label(text='Cotação de 1 Moeda Específica', borderwidth=2, relief='solid')
currency_quotation.grid(row=0, column=0, padx=10, pady=10, sticky='nswe', columnspan=3)

select_currency = tk.Label(text='Selecione a Moeda', anchor='e')
select_currency.grid(row=1, column=0, padx=10, pady=10, sticky='nswe', columnspan=2)

combobox_selectcurrency = ttk.Combobox(values=currency_list)
combobox_selectcurrency.grid(row=1, column=2, padx=10, pady=10, sticky='nswe')

select_day = tk.Label(text='Selecione o dia em que deseja pegar a Cotação', anchor='e')
select_day.grid(row=2, column=0, padx=10, pady=10, sticky='nswe', columnspan=2)

currency_calendar = DateEntry(year=2023, locale='pt_br')
currency_calendar.grid(row=2, column=2, padx=10, pady=10, sticky='nswe')

currency_text = tk.Label(text="")
currency_text.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky='nswe')

button_getcurrency = tk.Button(text='Pegar Cotação', command=pegar_cotacao)
button_getcurrency.grid(row=3, column=2, padx=10, pady=10, sticky='nswe')

# Cotação de varias moedas

quote_multiple_currency = tk.Label(text='Cotação de Múltiplas Moedas', borderwidth=2, relief='solid')
quote_multiple_currency.grid(row=4, column=0, padx=10, pady=10, sticky='nswe', columnspan=3)

select_file = tk.Label(text='Selecione um arquivo em Excel com as Moedas na coluna A: ')
select_file.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky='nswe')

var_pathfile = tk.StringVar()

button_selectfile = tk.Button(text='Clique para Selecionar', command=selecionar_arquivo)
button_selectfile.grid(row=5, column=2, padx=10, pady=10, sticky='nswe')

selected_file = tk.Label(text='Nenhum Arquivo Selecionado', anchor='e')
selected_file.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky='nswe')

initial_date = tk.Label(text='Data Inicial', anchor='e')
initial_date.grid(row=7, column=0, padx=10, pady=10, sticky='nswe')

end_date = tk.Label(text='Data Final', anchor='e')
end_date.grid(row=8, column=0, padx=10, pady=10, sticky='nswe')

calendar_start_date = DateEntry(year=2023, locale='pt_br')
calendar_start_date.grid(row=7, column=1, pady=10, padx=10, sticky='nsew')

calendar_end_date = DateEntry(year=2023, locale='pt_br')
calendar_end_date.grid(row=8, column=1, pady=10, padx=10, sticky='nsew')

update_quotation = tk.Button(text='Atualizar Cotações', command=atualizar_cotacoes)
update_quotation.grid(row=9, column=0, padx=10, pady=10, sticky='nsew')

text_update_quotation = tk.Label(text="")
text_update_quotation.grid(row=9, column=1, columnspan=2, padx=10, pady=10, sticky='nsew')

close_button = tk.Button(text='Fechar', command=window.quit)
close_button.grid(row=10, column=2, padx=10, pady=10, sticky='nswe')

window.mainloop()
