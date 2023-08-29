#Importações
import pandas as pd
import matplotlib.pyplot as plt
import datetime as datetime
from datetime import timedelta
import numpy as np
import locale
import win32com.client as win32
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import glob  

#Criando o navegador com a opção Headless
options = webdriver.ChromeOptions()
options.add_argument("--headless=chrome")
options.add_argument("--window-size=1920,1080")
navegador = webdriver.Chrome(options=options)
downloadPath = r'C:\Users\CarlosColmenero\Downloads'
params = {'behavior': 'allow', 'downloadPath': downloadPath}
navegador.execute_cdp_cmd('Page.setDownloadBehavior', params)

#Acessando o sistema
navegador.get('https://tmt.multitms.com.br/Login')
navegador.maximize_window()
navegador.find_element(By.ID, 'Usuario').send_keys('Usuario')
navegador.find_element(By.ID, 'Senha').send_keys('Senha', Keys.ENTER)
time.sleep(2)

#Acessando tela de relatórios
navegador.get(r'https://tmt.multitms.com.br/*')
time.sleep(2)

#Inicializando variáveis de data
inicio_mes = datetime.date.today().replace(day=1)
dia_anterior = datetime.date.today() - timedelta(days=1)
inicio_mes_anterior = inicio_mes.replace(month=inicio_mes.month-1)
fim_mes_anterior = inicio_mes - timedelta(days=1)

#Preenchendo filtros do relatório
navegador.find_element(By.XPATH, '/html/body/div[2]/div/div/main/div[2]/div/div[2]/form/div/div/div[1]/div/div/div/button').click()
time.sleep(1)
navegador.find_element(By.XPATH, '/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/form/div[2]/div/div/input').send_keys('Análise GAP', Keys.ENTER)
time.sleep(1)
navegador.find_element(By.XPATH, '/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/div[3]/div[1]/div/div/table/tbody/tr/td[3]/a').click()
navegador.find_element(By.XPATH, '/html/body/div[2]/div/div/main/div[2]/div/div[2]/form/div/div/div[4]/div/div/input').send_keys('01/{:02d}/{}'.format(inicio_mes_anterior.month, inicio_mes_anterior.year), Keys.TAB)
navegador.find_element(By.XPATH, '/html/body/div[2]/div/div/main/div[2]/div/div[2]/form/div/div/div[5]/div/div/input').send_keys('{}/{:02d}/{}'.format(dia_anterior.day, dia_anterior.month, dia_anterior.year), Keys.TAB)
navegador.find_element(By.XPATH, '//*[@id="knockoutPesquisaFreteTerceirizado"]/div/div[9]/div/div/button').click()
navegador.find_element(By.XPATH, '//*[@id="bs-select-1"]/ul/li[1]').click()
navegador.find_element(By.XPATH, '//*[@id="bs-select-1"]/ul/li[6]').click()
navegador.find_element(By.XPATH, '//*[@id="knockoutPesquisaFreteTerceirizado"]/div/div[9]/div/div/button').click()
navegador.find_element(By.XPATH, '/html/body/div[2]/div/div/main/div[2]/div/div[2]/div[1]/div[2]/button[2]').click()
navegador.execute_script("window.scrollBy(0,1000)","")
time.sleep(5)
navegador.find_element(By.XPATH, '/html/body/div[2]/div/div/main/div[2]/div/div[2]/div[4]/button[1]').click()


#Abrindo o arquivo base para tratamento
arquivo = glob.glob(r'C:\Users\CarlosColmenero\Downloads\Relatório_de_Fretes_Terceirizados_*.xls')

while len(arquivo) == 0:
    arquivo = glob.glob(r'C:\Users\CarlosColmenero\Downloads\Relatório_de_Fretes_Terceirizados_*.xls')
    time.sleep(1)
else: 
    relatorio_total_viagens_df = pd.read_excel(arquivo[0])

#analistando os tipos de dados do relatório
# relatorio_total_viagens_df.dtypes

#Separando o relatório: Viagens x Adicionais
relatorio_adicionais_df = relatorio_total_viagens_df[relatorio_total_viagens_df['Tipo documento'] == "MIN"]
relatorio_viagens_df = relatorio_total_viagens_df[relatorio_total_viagens_df['Tipo documento'] != "MIN"]

#Adicionando coluna de GAP
relatorio_viagens_df['gap'] = (1 - (relatorio_viagens_df['Vl. Bruto'] / relatorio_viagens_df['Vl. do Frete'])) * 100

#Exibindo DataFrames
# display(relatorio_adicionais_df)
# display(relatorio_viagens_df)

#Removendo veículos agregados - Possuem regra diferente de precificação
lista_agregados = ['GOV5046, END9988','MIR5039, CPI5614', 'BTA1528, CPI5631', 'CZC3050, CPI6156', 'ELW4H92, CPG2759', 'JOD8I15, MFX9988', 'JLD3B46,  CPI5614']
relatorio_viagens_sem_agregados_df = relatorio_viagens_df[~relatorio_viagens_df['Veículo'].isin(lista_agregados)]
# display(relatorio_viagens_sem_agregados_df)  

#Salvando base
relatorio_viagens_sem_agregados_df.to_excel(r'C:\Users\CarlosColmenero\Downloads\base_viagens.xlsx', index=False)

#Editando datas para formato Numpy v.
locale.setlocale(locale.LC_TIME, 'pt_BR.UFT-8')
inicio_mes = np.datetime64(inicio_mes)
dia_anterior = np.datetime64(dia_anterior)
inicio_mes_anterior = np.datetime64(inicio_mes_anterior)
fim_mes_anterior = np.datetime64(fim_mes_anterior)

#Removendo inconsistências do DataFrame
relatorio_viagens_sem_agregados_df = relatorio_viagens_sem_agregados_df.loc[(relatorio_viagens_sem_agregados_df['gap'] >= 0), :]

#Filtrando DataFrame por período
relatorio_mes_anterior = relatorio_viagens_sem_agregados_df.loc[(relatorio_viagens_sem_agregados_df['Data Emissão'] >= inicio_mes_anterior) & (relatorio_viagens_sem_agregados_df['Data Emissão'] <= fim_mes_anterior), :]
relatorio_mes_atual = relatorio_viagens_sem_agregados_df.loc[(relatorio_viagens_sem_agregados_df['Data Emissão'] >= inicio_mes), :]
relatorio_dia_anterior = relatorio_viagens_sem_agregados_df.loc[(relatorio_viagens_sem_agregados_df['Data Emissão'] >= dia_anterior), :]

#Agrupando indicado por tipo de operação
relatorio_mes_anterior = relatorio_mes_anterior[['Tipo Operação','gap']].groupby('Tipo Operação',as_index=False).mean()
relatorio_mes_atual = relatorio_mes_atual[['Tipo Operação','gap']].groupby('Tipo Operação', as_index=False).mean()
relatorio_dia_anterior = relatorio_dia_anterior[['Tipo Operação','gap']].groupby('Tipo Operação', as_index=False).mean()

#Exibindo relatórios filtrados e agrupados
# print('Mes Anterior')
# display(relatorio_mes_anterior)
# print('-'*60)
# print('Mes Atual')
# display(relatorio_mes_atual)
# print('-'*60)
# print('Ontem')
# display(relatorio_dia_anterior)

#Unindo relatórios em um único DataFrame
relatorio_gap_df = relatorio_mes_anterior.merge(relatorio_mes_atual, on=['Tipo Operação'], how="outer").merge(relatorio_dia_anterior, on=['Tipo Operação'], how="outer")

#Renomeando as colunas
mes_anterior_rotulo = '{}'.format((datetime.date.today().replace(day=1) - timedelta(days=1)).strftime('%b/%Y'))
mes_atual_rotulo = '{}'.format(datetime.date.today().strftime('%b/%Y'))
ultimo_dia_rotulo = '{}/{}'.format((datetime.date.today() - timedelta(days=1)).day, (datetime.date.today() - timedelta(days=1)).strftime('%b'))
relatorio_gap_df = relatorio_gap_df.rename(columns={"gap_x": mes_anterior_rotulo, "gap_y": mes_atual_rotulo, 'gap': ultimo_dia_rotulo})
relatorio_gap_df = relatorio_gap_df.fillna(0)
# display(relatorio_gap_df)

#Editando formato dos números
pd.options.display.float_format = '{:,.2f}%'.format

#Criando indicadores visuais
condicao = (relatorio_gap_df[mes_atual_rotulo] == 0) | (relatorio_gap_df[mes_anterior_rotulo] == 0)
relatorio_gap_df['Variacao Mensal'] = np.where(condicao, '-', (np.where(relatorio_gap_df[mes_atual_rotulo] > relatorio_gap_df[mes_anterior_rotulo], '🟢', '🔴')))
relatorio_gap_df['Variacao Diária'] = np.where(relatorio_gap_df[ultimo_dia_rotulo] == 0, '-', (np.where(relatorio_gap_df[ultimo_dia_rotulo] > relatorio_gap_df[mes_atual_rotulo], '🟢', '🔴')))

#Filtrando relatório Final
relatorio_gap_df = relatorio_gap_df[['Tipo Operação', mes_anterior_rotulo, mes_atual_rotulo, 'Variacao Mensal', ultimo_dia_rotulo, 'Variacao Diária']]
# display(relatorio_gap_df)

#Filtrando lista para o gráfico
lista_operacoes = ['NESTLÉ', 'PURINA', 'LEROY MERLIN', 'LM - REVERSA', 'UNILEVER', 'P&G']
base_grafico = relatorio_gap_df[relatorio_gap_df['Tipo Operação'].isin(lista_operacoes)]
# display(base_grafico)    

#Parametrizando a figura
plt.figure(figsize= (30, 15))
barWidth = 0.25

#Altura das barras
bars1 = base_grafico[mes_anterior_rotulo].tolist()
bars2 = base_grafico[mes_atual_rotulo].tolist()
# bars3 = base_grafico[ultimo_dia_rotulo].tolist()
 
#Posição do eixo x
r1 = np.arange(len(bars1))
r2 = [x + barWidth for x in r1]
# r3 = [x + barWidth for x in r2]
 
#Criando o gráfico
barra1 = plt.bar(r1, bars1, width=barWidth, label=mes_anterior_rotulo, color='darkslategray')
plt.bar_label(barra1, label=bars1, padding=3, fontsize=20, fmt='%.2f') 
barra2 = plt.bar(r2, bars2, width=barWidth, label=mes_atual_rotulo, color='lightseagreen')
plt.bar_label(barra2, label=bars2, padding=3, fontsize=20, fmt='%.2f')
# barra3 = plt.bar(r3, bars3, width=barWidth, edgecolor='white', label=ultimo_dia_rotulo)
# plt.bar_label(barra3, label=bars3, padding=3, fontsize=16, fontweight='bold', fmt='%.2f')

#Adicionando ticks
plt.title(f'Evolução GAP {mes_anterior_rotulo} x {mes_atual_rotulo}', fontweight='bold', fontsize=30)
plt.xticks([r + barWidth/2 for r in range(len(bars1))], base_grafico['Tipo Operação'].tolist(), rotation='vertical', fontsize=20, fontweight='bold')
plt.yticks([])

#Criando Legendas e apresentando o gráfico
plt.legend(fontsize='xx-large', frameon=False, loc='upper left')
plt.tick_params(axis='x', length=0)
plt.box(False)
plt.savefig(r'C:\Users\CarlosColmenero\Downloads\grafico.png', bbox_inches='tight', dpi=40)
# plt.show()

#Parametrizando o envio de e-mail
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'lista@lista.com'
# mail.CC = 
mail.Subject = f'Análise de GAP - {ultimo_dia_rotulo}'
#mail.Body = 
mail.HTMLBody = f"""
Prezados, bom dia,
<html><body><br></body></html>
Segue análise de GAP das operações para os meses de {mes_anterior_rotulo} e {mes_atual_rotulo}, a base está atualizada com as vianges até o dia {ultimo_dia_rotulo}.
<html><body><br></body></html>
A tabela abaixo trás o resumo e no anexo temos a base total de viagens.
<html><body><br></body></html>
Cabe ressaltar que não estão contemplados os fretes realizados pelos agregados e que o GAP é calculado por: Frete Bruto Terceiro \ Frete CTe Líquido de ICMS.
<html><body><br></body></html>
    {relatorio_gap_df.to_html(index=False, decimal='.')}
<html><body><br></body></html>
    <html><body><img src="C:\\Users\\CarlosColmenero\\Downloads\\grafico.png" style="width:100%"/></p></body></html>
<html><body><br></body></html>
Em caso de dúvidas fico à disposição.
<html><body><br></body></html>
Atenciosamente,
<html><body><img src='C:\\Users\\CarlosColmenero\\TMTLOG\\TMT CORPORATIVO - Documentos\\TMTLOG\\Carlos\\assinatura_email.png' style="width:100%"/></p></body></html>
"""    
#Anexos:
attachment  = r'C:\\Users\\CarlosColmenero\\Downloads\\base_viagens.xlsx'
mail.Attachments.Add(attachment)
mail.Send()





