import smtplib
import email.message
import pandas as pd
from bs4 import BeautifulSoup
import os
import time
import random

excel_file = 'planilha-teste.xlsx'
sheet_name = 'Planilha1'

tempo_min = 60
tempo_max = 75

tempo_aleatorio = random.uniform(tempo_min, tempo_max)

def send_email():
    body_email = f"""
    <p>Bom Dia!</p>
    <p>Sua solicitação ainda se encontra em situação pendente, por gentileza verificar.</p>
    <p>Caso já foi verificado ignore essa mensagem. </p>
    <p>{full_html}</p>
    <p><img src="https://e7.pngegg.com/pngimages/391/392/png-clipart-jpeg-signature-scalable-graphics-sal-atilde-o-angle-text.png"></p>
    """
    msg = email.message.Message()
    msg['Subject'] = f"PROPOSTA REFERENTE AO BANCO X | ID {id_value}"
    #coloque o e-mail de envio
    msg['From'] = " " 
    msg['To'] = f"{email_gerente}, {email_parceiro}"
    #gere uma senha de app nas configuracões de e-mail
    senha_from = " "
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(body_email)
    s = smtplib.SMTP("smtp.gmail.com: 587")
    s.starttls()
    s.login(msg['From'], senha_from)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print(f'Email enviado com sucesso para o id {id_value}')
    time.sleep(tempo_aleatorio)

df = pd.read_excel(excel_file, sheet_name=sheet_name)

def dataframe_to_html(df):
    df_filtered = df.drop(columns=['e-mail1', 'e-mail2'], errors='ignore')
    return df_filtered.to_html(index=False, classes='my-table')

unique_ids = df['ID'].unique()
print(f'IDs encontrados: {unique_ids}')

for id_value in unique_ids:
    group = df[df['ID'] == id_value]
    if group.empty:
        print(f'Grupo vazio para ID: {id_value}')
        continue
    html_table = dataframe_to_html(group)
    email_parceiro = group['e-mail1'].iloc[0]  
    email_gerente = group['e-mail2'].iloc[0]  
    html_head = """
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            .my-table {
                width: 100%;
                border-collapse: collapse;
            }
            .my-table th, .my-table td {
                border: 1px solid black;
                padding: 8px;
                text-align: center;
            }
            .my-table th {
                background-color: #080808; 
                color: #fff; 
            }
            .my-table td {
                background-color: #131313; 
                color: #fff;
            }
        </style>
    </head>
    <body>
    """
    html_end = "</body></html>"
    soup = BeautifulSoup(html_table, 'html.parser')
    table = soup.find('table')
    for th in table.find_all('th'):
        th['style'] = 'background-color: #131313; color: #fff;'
    for td in table.find_all('td'):
        td['style'] = 'background-color: #ffffff; color: #000000;'

    full_html = html_head + str(soup) + html_end
    send_email()

print("envios realizados.")