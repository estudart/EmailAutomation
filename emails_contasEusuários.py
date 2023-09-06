import pandas as pd
from datetime import datetime
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import (Mail, From)
import time

SENDGRID_KEY = '<api_key'
sg = SendGridAPIClient(SENDGRID_KEY)
sender = '<email>'

# Planilha que contém as contas para cada corretora
accounts_register = pd.read_excel("P:\Python ETS\Cadastro de contas e usuarios\planilha_envio.xlsx")

# Planilha que contém os reponsáveis por cadastrar as contas
accounts_email = pd.read_excel("P:\Python ETS\Cadastro de contas e usuarios\cadastro_responsaveis.xlsx")


# Função que cria o email
def create_email(broker, name, request_type, receive_email, account, institution, call, client_email):

    if account != "":
        accounts = "<br />".join(account)

    if len(name) == 1:
        is_many_clients = False
        names = f'{name[0]}'
    else:
        is_many_clients = True
        names = f"{', '.join([client for client in name[:-1]])} e {name[-1]}"

    now = datetime.now()
    is_morning = now.hour < 12
    is_night = now.hour > 18

    if is_morning:
        greeting = "bom dia"
    elif is_night:
        greeting = "boa noite"
    else:
        greeting = "boa tarde"

    if account != "":
        if broker == "CREDIT SUISSE":
            email_body = f"""
            <span style='font-family:verdana;font-size:12.0;color:#1F497D'>Prezados, {greeting}. <br><br>
            
            {names}, da instituição {institution},
            {'estão' if is_many_clients else 'está'} solicitando acesso para operar na(s) conta(s) abaixo:<br><br>
            {accounts} <br><br>
            
            Podemos autorizar? <br><br>
            
            {names}, {'poderiam' if is_many_clients else 'poderia'} confirmar a solicitação? <br><br>
            
            Caso positivo, devemos validar limite na ATG? <br><br>
            
            Atenciosamente, <br><br>
            """

        else:
            email_body = f"""
            <span style='font-family:verdana;font-size:12.0;color:#1F497D'>Prezados, {greeting}. <br><br>

            {names}, da instituição {institution},
            {'estão' if is_many_clients else 'está'} solicitando acesso para operar na(s) conta(s) abaixo:<br><br>
            {accounts} <br><br>

            Podemos autorizar? <br><br>

            Caso positivo, devemos validar limite na ATG? <br><br>

            Atenciosamente, <br><br>
            """
    else:
        if broker == "CREDIT SUISSE":
            email_body = f"""
            <span style='font-family:verdana;font-size:12.0;color:#1F497D'>Prezados, {greeting}. <br><br>

            {names}, da instituição {institution},
            {'estão' if is_many_clients else 'está'} solicitando acesso para operar na(s) conta(s) abaixo:<br><br>
            {accounts} <br><br>

            Podemos autorizar? <br><br>
            
            {names}, {'poderiam' if is_many_clients else 'poderia'} confirmar a solicitação? <br><br>

            Caso positivo, devemos validar limite na ATG? <br><br>

            Atenciosamente, <br><br>
            """
        else:
            email_body = f"""
            <span style='font-family:verdana;font-size:12.0;color:#1F497D'>Prezados, {greeting}. <br><br>

            {names}, da instituição {institution},
            {'estão' if is_many_clients else 'está'} solicitando acesso para operar na(s) conta(s) abaixo:<br><br>
            {accounts} <br><br>

            Podemos autorizar? <br><br>

            Caso positivo, devemos validar limite na ATG? <br><br>

            Atenciosamente, <br><br>
            """

    subject = f"{institution} > {broker} | Cadastro de {request_type} - #{call}"

    message = Mail(
        from_email=From(sender, 'ATG | Electronic Trading Support'),
        to_emails=receive_email,
        subject=subject,
        html_content=email_body
    )

    message.add_cc('servicedesk@americastg.com')
    message.add_cc('ets@americastg.com')

    if broker == "CREDIT SUISSE":
        for i in client_email:
            message.add_cc(str(i))

    return message


institution = accounts_register.loc[0, 'institution']
client_email = accounts_register.loc[0, 'email_cliente'].split(";")
call = str(accounts_register.loc[0, 'chamado'])
brokers = accounts_register['broker'].tolist()
broker_email = accounts_email['broker'].tolist()


def send_email(message):
    sg.send(message)
    print(message.subject)


for k in range(len(brokers)):

    if type(brokers[k]) == float:
        break

    request_type = accounts_register.loc[k, 'request']
    name = accounts_register.loc[k, 'name'].split(",")
    # print(name)

    if type(accounts_register.loc[k, 'account_list']) is not str:
        account = ""
    else:
        account = accounts_register.loc[k, 'account_list'].split(',')

    broker = brokers[k]
    receive_emails = accounts_email[accounts_email.broker == brokers[k]].to_dict(orient="records")[0]

    if request_type == 'conta':
        receive_email = receive_emails['account_email_list'].split(";")
    else:
        receive_email = receive_emails['user_email_list'].split(";")

    message = create_email(broker, name, request_type, receive_email, account, institution, call, client_email)
    # print(message)
    send_email(message)
    time.sleep(3)
