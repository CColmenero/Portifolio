#!/usr/bin/env python
# coding: utf-8

# In[4]:


import requests
import json
from tkinter import *
import tkinter.filedialog
import pandas as pd
import win32api
from winreg import *

#Abrindo caixa para seleção de arquivo
root= Tk()
arquivo = tkinter.filedialog.askopenfilename(title = "Selecione o Arquivo csv com Canais e Keywords")
root.destroy()

#Chave de API
api_key = 'xxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxx-xxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxxxx'

#Definindo fução para consulta via API CNNJ Já
def get_company(tax_id):
  url = "https://api.cnpja.com/office/" + tax_id + "?simples=true"
  headers = { 'Authorization': api_key }
  response = requests.request("GET", url, headers = headers)
  return json.loads(response.text)

#Lendo base
base = pd.read_excel(arquivo)
base = base['CPF/CNPJ Terceiro'].unique()

#Looping para tratamento e consulta de base
lista = []

for cnpj in base:
    cnpj = cnpj.replace('.', '').replace('/', '').replace('-','')
    if len(cnpj) == 14:
        tax_id = cnpj
        company = get_company(tax_id)
        json_str = json.dumps(company)
        resp = json.loads(json_str)
        lista.append((cnpj, resp['company']['simples']['optant']))

#Encontrando pasta de Downloads
with OpenKey(HKEY_CURRENT_USER, 'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders') as key:
    downloads = QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')[0]
        
        
#Extraído base com informação da situação        
df = pd.DataFrame(lista, columns=['cnpj', 'situação'])
df['situação'] = df['situação'].replace(True, 'Simples')
df['situação'] = df['situação'].replace(False, 'Não Simples')
df.to_excel(downloads+r'\situacao_PJ.xlsx', index=False)

#Exbindo mensagem de final de consulta
win32api.MessageBox(0, 'A consulta acabou, arquivo disponível na pasta de downloads', 'Consulta Simples Nacional')

