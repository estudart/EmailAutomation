from datetime import datetime
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import (Mail, From, Attachment, FileContent, FileName, FileType, Disposition)
import pandas as pd
import os
import base64
from unidecode import unidecode
import time

# API KEY SENDGRID

SENDGRID_KEY = '<api_key'
sg = SendGridAPIClient(SENDGRID_KEY)
sender = '<email>'

dict_month = {
    '01': ['01', '1', 'Janeiro', 'jan', 'JANEIRO'],
    '02': ['02', '2', 'Fevereiro', 'fev', 'FEVEREIRO'],
    '03': ['03', '3', 'Março', 'mar', 'MARÇO'],
    '04': ['04', '4', 'Abril', 'abr', 'ABRIL'],
    '05': ['05', '5', 'Maio', 'mai', 'MAIO'],
    '06': ['06', '6', 'Junho', 'jun', 'JUNHO'],
    '07': ['07', '7', 'Julho', 'jul', 'JULHO'],
    '08': ['08', '8', 'Agosto', 'ago', 'AGOSTO'],
    '09': ['09', '9', 'Setembro', 'set', 'SETEMBRO'],
    '10': ['10', '10', 'Outubro', 'out', 'OUTUBRO'],
    '11': ['11', '11', 'Novembro', 'nov', 'NOVEMBRO'],
    '12': ['12', '12', 'Dezembro', 'dez', 'DEZEMBRO']
    }

dict_broker = {
                'C6': ['C6'],
                'ITAU': ['ITAU'],
                'ATIVA': ['ATIVA'],
                'BTG': ['BTG'],
                'CM': ['CM'],
                'CREDIT': ['CREDITSUISSE'],
                'GENIAL': ['GENIAL'],
                'GOLDMAN': ['GOLDMANSACHSDOBRASILCTVM'],
                'GUIDE': ['GUID'],
                'IDEAL': ['IDEAL'],
                'INTER': ['INTER'],
                'INTL': ['FCSTONE'],
                'JP': ['JPMORGANCCVMSA'],
                'MERRILL': ['ML'],
                'MIRAE': ['MIRAE'],
                'MODAL': ['MODAL'],
                'MORGAN': ['MS'],
                'NECTON': ['NECT'],
                'NOVA': ['NOVAFUTURA'],
                'RENASCENCA': ['RENA'],
                'TERRA': ['TERRA'],
                'TULLETT': ['TULLETT'],
                'UBS': ['UBSBB'],
                'VITREO': ['VITREO'],
                'XP': ['XP'],
                'BGC': ['BGC'],
                'BRADESCO': ['BRAD'],
                'ORAMA': ['ORAMA'],
                'WARREN': ['WARREN'],
                'CITI': ['CITI'],
                'SAFRA': ['SAFRA'],
                'SANTANDER': ['SANT'],
                'PLANNER': ['PLAN'],
                'LEV': ['LEVDTVM'],
                'RB': ['RBINVESTIMENTOSDTVM'],
                'BB': ['BBINVESTIMENTOSA']
                }


# Função que dispara os e-mails
def send_email(message):
    sg.send(message)
    print(message.subject)


# Função que cria o email a ser enviado
def create_email(broker, broker_row, attachments, path, month, year, call):
    attachment_atg, attachment_broker = attachments

    now = datetime.now()
    is_morning = now.hour < 12
    is_night = now.hour > 18
    if is_morning:
        greeting = "bom dia"
    elif is_night:
        greeting = "boa noite"
    else:
        greeting = "boa tarde"

    receive_email = broker_row["email_list"].split(";")

    if broker_row['type'] == 'autonomo':
        email_body = f"""
        Prezado(a)s, {greeting}. <br><br>
        Segue em anexo relatório de Latência e Disponibilidade + Usuários de Market Data. <br><br>
        {broker}_{month}-{year}.csv(Lista de Usuários) <br>
        {broker}_{month}-{year}_BOV.csv(Report que deverá ser incluído no Report da Bolsa para BOVESPA) <br>
        {broker}_{month}-{year}_BMF.csv(Report que deverá ser incluído no Report da Bolsa para BMF) <br><br>
        *No relatório de usuários anexo({broker}_{month}-{year}) estão incluídos usuários do tipo DMA que consumiram MKT Data,
        os mesmos não devem ser reportados pela corretora, pois já estão sendo reportados pela ATG, conforme
        política de MKT Data da B3. Sendo assim, a corretora deve reportar somente os usuários da mesa."""
    else:
        email_body = f"""
        Prezado(a)s, {greeting}. <br><br>
        Segue em anexo relatório de Latência e Disponibilidade + Usuários de Market Data. <br><br>
        Qualquer dúvida, estou à disposição. <br><br>
        Atenciosamente,"""

    if broker != 'CITI':
        subject = f"{broker} | Relatório de Latência e Disponibilidade + Usuários Mkt {month}-{year} - #{call}"
    else:
        subject = f"Brazil Tech Third Party Management: Service Level Agreement - SLA - AMERICAS TRADING GROUP (ATG) #{call}"

    message = Mail(
        from_email=From(sender, 'ATG | Electronic Trading Support'),
        to_emails=receive_email,
        subject=subject,
        html_content=email_body
    )
    message.add_cc('servicedesk@americastg.com')
    message.add_cc('ets@americastg.com')
    message.attachment = [attachment_atg, attachment_broker]

    if broker_row['type'] == 'autonomo':
        attachment_broker_auto = create_attachment(
            att_path=path + f'RELATORIO DE MKT/{year}/{dict_month[month][0]} - {dict_month[month][4]}/atg-to-brokers/AUTONOMOUS',
            att_file=f'{broker}_{dict_month[month][0]}-{year}.csv',
            filetype='application/pdf'
        )

        attachment_broker_auto_1 = create_attachment(
            att_path=path + f'RELATORIO DE MKT/{year}/{dict_month[month][0]} - {dict_month[month][4]}/brokers-to-b3',
            att_file=f'{broker}_{dict_month[month][0]}-{year}_BMF.csv',
            filetype='application/pdf'
        )

        attachment_broker_auto_2 = create_attachment(
            att_path=path + f'RELATORIO DE MKT/{year}/{dict_month[month][0]} - {dict_month[month][4]}/brokers-to-b3',
            att_file=f'{broker}_{dict_month[month][0]}-{year}_BOV.csv',
            filetype='application/pdf'
        )

        message.add_attachment(attachment_broker_auto)
        message.add_attachment(attachment_broker_auto_1)
        message.add_attachment(attachment_broker_auto_2)

    elif broker_row['type'] == 'dependente':

        attachment_broker_dep = create_attachment(
            att_path=path + f'RELATORIO DE MKT/{year}/{dict_month[month][0]} - {dict_month[month][4]}/atg-to-brokers/DEPENDENT/',
            att_file=f'{broker}_{dict_month[month][0]}-{year}.csv',
            filetype='application/pdf'
        )
        message.add_attachment(attachment_broker_dep)

        if broker == 'SAFRA':
            attachment_broker_safra1 = create_attachment(
                att_path=path + f'AUDITORIAS/INDICADORES DE LATÊNCIA E DISPONIBILIDADE/{year}/{dict_month[month][0]} - {dict_month[month][4]}',
                att_file=f'{dict_month[month][0]}-{year} Acompanhamento de Indicadores de Performance - ATG - {broker}.XLSX',
                filetype='application/pdf'
            )
            message.add_attachment(attachment_broker_safra1)

            attachment_broker_safra2 = create_attachment(
                att_path=path + f'AUDITORIAS/CLIENTES/SAFRA - 59/{year}',
                att_file=f'User by broker_{broker}_{day}-{dict_month[month_real][0]}-{year}.xlsx',
                filetype='application/pdf'
            )
            message.add_attachment(attachment_broker_safra2)

        if broker == 'SANTANDER':
            attachment_broker_sant1 = create_attachment(
                att_path=path + f'AUDITORIAS/INDICADORES DE LATÊNCIA E DISPONIBILIDADE/{year}/{dict_month[month][0]} - {dict_month[month][4]}',
                att_file=f'Relatório de latências - {dict_month[month][2]}{year} - {broker}.pdf',
                filetype='application/pdf'
            )
            message.add_attachment(attachment_broker_sant1)

    return message


def create_attachment(att_path, att_file, filetype):
    filename = att_file
    att_file = f"{att_path}/{att_file}"

    # checa o caminho
    if not os.path.exists(att_file):
        print(f"Caminho não existe: {att_file}")
        return
    # le o arquivo
    with open(att_file, 'rb') as f:
        data = f.read()
        f.close()

    # codifica o conteudo do arquivo
    encoded_file = base64.b64encode(data).decode()

    # cria o anexo
    return Attachment(
        FileContent(encoded_file),
        FileName(unidecode(filename)),
        FileType(filetype),
        Disposition('attachment')
    )

##################

# Inicio do script


base_path = 'Z:/Internal/'
emails_df = pd.read_excel("Z:\Internal\PYTHON ETS\email_latencia.xlsx")

# Pergunta que inputa a data
now = datetime.now()
date = now.strftime("%d/%m/%Y")

# date = '21/09/2022'
# chamado = 'SD-34251'
year = date[6:]
month = '12' if date[3:5] == '01' else f"{int(date[3:5]) - 1:02}"
#import pdb; pdb.set_trace()# joga pra dezembro se janeiro for informado
month_real = date[3:5]
day = date[:2]

# Pergunta que inputa o chamado
chamado = str(emails_df.loc[0, 'chamado'])

brokers = emails_df['broker'].tolist()
# type_broker = emails_df['type'].tolist()
not_sent_emails = ['nenhum']

for broker in brokers:

    attachment_atg = create_attachment(
        att_path=base_path + f'AUDITORIAS/DISPONIBILIDADE ATG/{year}/{dict_month[month][0]} - {dict_month[month][4]}',
        att_file=f'{dict_month[month][0]}-ATG - Relatório de Disponibilidade - {dict_month[month][2]} - {year}.pdf',
        filetype='application/pdf')

    attachment_broker = create_attachment(
        att_path=base_path + f'AUDITORIAS/INDICADORES DE LATÊNCIA E DISPONIBILIDADE/{year}/{dict_month[month][0]} - {dict_month[month][4]}',
        att_file=f'Indicadores_de_Latencia_e_Disponibilidade_{dict_month[month][3]}-{year}-{dict_broker[broker][0]}.pdf',
        filetype='application/pdf')

    try:
        broker_row = emails_df[emails_df.broker == broker].to_dict(orient="records")[0]

        message = create_email(
            broker=broker,
            broker_row=broker_row,
            attachments=(attachment_atg, attachment_broker),
            path=base_path,
            month=month,
            year=year,
            call=chamado
            )
        send_email(message)
        time.sleep(3)

    except Exception as err:
        not_sent_emails.append(broker_row["broker"])


if len(not_sent_emails) > 0:
    print(f"Não foi possível enviar email para os seguintes destinatários: {not_sent_emails}")

