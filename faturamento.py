from datetime import datetime
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import (Mail, From, Attachment, FileContent, FileName, FileType, Disposition)
import pandas as pd
import os
import base64
from unidecode import unidecode
from tkinter import messagebox, ttk, Label, Button
import tkinter as tk

SENDGRID_KEY = '<api_key>'
sg = SendGridAPIClient(SENDGRID_KEY)
sender = '<email>'

# emails_df = pd.read_excel("Z:\Controladoria\Faturamento\1. ATG\Clientes\Cadastro_emails_ATG.xlsx")

path = "Z:/Controladoria/Faturamento"

# Planilha que contém as informações para cada corretora
registros_df = pd.read_excel(f"{path}/cadastro_corretoras.xlsx", sheet_name='completo')


# Função que cria o email a ser enviado
def create_email(broker, broker_row, month, year, path):
    global subject
    global receive_email
    global email_body
    global message
    #attachment_atg, attachment_broker = attachments

    lista = []
    lista2 = []

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
    receive_email_screen = str("\n".join(receive_email))
    nf = broker_row["nf"]

    if broker_row['type'] == 'padrao':
        email_body = f"""
        <span style='font-family:verdana;font-size:12.0;color:#1F497D'>
Prezado(a)s, {greeting}. <br><br>
        
Segue anexo o faturamento referente aos serviços prestados. <br><br>
        
Abaixo os dados bancários para depósito: <br><br>
        
Enterprise S/A <br>
CNPJ XXX.XXX.XXX/0001-31 <br>
BANK (341) <br>
AG. XXXX CC. XXXX-X <br><br>
        
Peço a gentileza de confirmar o recebimento. <br><br>
        
Qualquer dúvida estamos à disposição. <br>
Atenciosamente,"""

    elif broker_row['type'] == 'boleto':
        email_body = f"""
        <span style='font-family:verdana;font-size:12.0;color:#1F497D'>
Prezado(a)s, {greeting}. <br><br>
        
Segue anexo do faturamento referente aos serviços prestados. <br><br>
        
Peço a gentileza de confirmar o recebimento.<br><br>
        
Qualquer dúvida, estou à disposição. <br><br>
        
Atenciosamente,"""

        attachment_broker_boleto = create_attachment(
            att_path=path,
            att_file=f'{year}-{month}-BOLETO-{broker}.pdf',
            filetype='application/pdf'
        )
        lista.append(attachment_broker_boleto)
        lista2.append(f'{year}-{month}-BOLETO-{broker}.pdf')

    else:
        email_body = f"""
        <span style='font-family:verdana;font-size:12.0;color:#1F497D'>
Dear Customer, <br><br>
        
Your ATG invoice is ready, please see attached. <br>
PLEASE WIRE FUNDS IN U.S. DOLLARS AND PROCEED INVOICE PAYMENT TO THE FOLLOWING ACCOUNT: <br><br>
        
Wire instructions (USD) <br>
Correspondent Bank: Standard Chartered Bank – New York – USA Swift <br>
(BIC CODE): XXXXXXX <br>
Clearing Code: ABA: XXXXX CHIPS UID: XXXX Account number: XXXXXXX <br>
Beneficiary Bank: Banco Santander (Brasil) S.A. <br>
Swift (BIC CODE): XXXXX <br>
Beneficiary Name: XXXXX-XXXX-XXXX Enterprise S/A <br>
Reference: Invoice Number <br><br>
        
** THIS LOCATION DOES NOT ACCEPT CHECKS **  <br><br>
        
Best regards,"""

    if broker == "J.P.MORGAN":
        attachment_broker_analitico = create_attachment(
            att_path=path,
            att_file=f'{year}-{month}-Relatorio Gerencial analitico-{broker}.xlsx',
            filetype='application/pdf'
        )
        lista.append(attachment_broker_analitico)
        lista2.append(f'{year}-{month}-Relatorio Gerencial analitico-{broker}.xlsx')

    subject = f"{broker} | Faturamento {month}-{year}"

    message = Mail(
        from_email=From(sender, 'ATG | Faturamento'),
        to_emails=receive_email,
        subject=subject,
        html_content=email_body
    )

    """attachment_broker_ordens = create_attachment(
        att_path=path,
        att_file=f'{year}-{month}-Relatorio-Ordens-{broker}.xlsx',
        filetype='application/pdf'
    )
    lista.append(attachment_broker_ordens)
    lista2.append(f'{year}-{month}-Relatorio-Ordens-{broker}.xlsx')"""

    attachment_broker_gerencial = create_attachment(
        att_path=path,
        att_file=f'{year}-{month}-Relatorio-Gerencial-{broker}.xlsx',
        filetype='application/pdf'
    )
    lista.append(attachment_broker_gerencial)
    lista2.append(f'{year}-{month}-Relatorio-Gerencial-{broker}.xlsx')

    attachment_broker_fatura = create_attachment(
        att_path=path,
        att_file=f'{year}-{month}-Fatura-{broker}-Roteamento.pdf',
        filetype='application/pdf'
    )
    lista.append(attachment_broker_fatura)
    lista2.append(f'{year}-{month}-Fatura-{broker}-Roteamento.pdf')

    for filename in os.listdir(path):
        f = os.path.join(path, filename)
        if os.path.isfile(f):
            if "NF" in f:
                file = f.replace('\\', "/")
                file3 = file.split("/")
                file3 = file3[:-1]
                file3 = "/".join(file3)
                print(file3)
                file2 = file.split("/")[-1]
                print(file2)
                attachment_broker_nota = create_attachment(
                    att_path=file3,
                    att_file=file2,
                    filetype='application/pdf'
                    )
                lista.append(attachment_broker_nota)
                lista2.append(str(file2))

    message.add_cc('pmachado@americastg.com')
    message.add_cc('aborghesan@americastg.com')
    message.add_cc('vcarrico@americastg.com')
    message.attachment = lista

    root = tk.Tk()
    root.title('Financeiro ATG')

    app_width = 800
    app_height = 700

    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    x = (screen_width / 2) - (app_width / 2)
    y = (screen_height / 2) - (app_height / 2)

    root.geometry(f"{app_width}x{app_height}+{int(x)}+{int(y)}")

    root.columnconfigure(0, weight=1)
    root.columnconfigure(1, weight=3)
    root.columnconfigure(2, weight=1)

    # L1 = Label(text='Financeiro ATG', font=40)
    # L1.grid(row=0, column=0)

    E1 = ttk.Label(root, text=subject, font=("Arial Black", 12), justify='center')
    E1.grid(column=1, row=0, sticky=tk.W, padx=5, pady=5)

    limpa_screen = ["""<span style='font-family:verdana;font-size:12.0;color:#1F497D'>"""]
    email_body_screen = email_body.replace(limpa_screen[0], " ")

    E2 = ttk.Label(root, text=str(email_body_screen.replace("<br>", "")), font=("Calibri", 12))
    E2.grid(column=1, row=2, sticky=tk.EW, padx=5, pady=5)

    E4 = ttk.Label(root, text=receive_email_screen, font=("Arial Black", 10))
    E4.grid(column=1, row=3, sticky=tk.EW, padx=5, pady=5)

    E5 = ttk.Label(root, text=str("\n".join(lista2)), font=("Calibri", 12))
    E5.grid(column=1, row=4, sticky=tk.EW, padx=5, pady=5)

    def click():
        label1 = Label(root, text='email enviado', font=("Arial Black", 10))
        label1.grid(column=1, row=6, sticky=tk.W, padx=5, pady=5)

    B1 = Button(root, text="Enviar", font=("Arial", 10), command=lambda: [click(), send()])
    B1.grid(column=1, row=5, sticky=tk.W, padx=5, pady=5)

    B2 = Button(root, text="Fechar", font=("Arial", 10), command=root.destroy)
    B2.grid(column=1, row=5, padx=5, pady=5)

    root.mainloop()

    print(f'******************************************* Fim do email *******************************************')
    return


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


now = datetime.now()
date = now.strftime("%d/%m/%Y")

year = date[6:]
month = '12' if date[3:5] == '01' else f"{int(date[3:5]) - 1:02}"

#import pdb; pdb.set_trace()# joga pra dezembro se janeiro for informado
month_real = date[3:5]
day = date[:2]


def send():
    global email_body_print
    global subject
    global receive_email
    global email_body
    global message
    print(f'{subject}')
    print(f'{receive_email}')
    limpa_span = ["""<span style='font-family:verdana;font-size:12.0;color:#1F497D'>"""]
    email_body_print = email_body.replace(limpa_span[0], " ")
    email_body_print = email_body_print.replace('<br>', "")
    print(f'{email_body_print}')
    print('enviou')
    sg.send(message)
    return


brokers_list = registros_df['broker'].tolist()
not_sent_emails = ['nenhum']


def run():
    broker_dropdown = combo.get()
    broker_row = registros_df[registros_df.broker == broker_dropdown].to_dict(orient="records")[0]
    path_arq = f"Z:/Controladoria/Faturamento/1. ATG/Clientes/{broker_dropdown}/{year}/{month}"
    create_email(broker_dropdown, broker_row, month, year, path_arq)


def encerrar():
    main_window.destroy()
    #root.destroy()


main_window = tk.Tk()
main_window.config(width=300, height=200)
main_window.title("Selecione o cliente")
combo = ttk.Combobox(
    state="readonly",
    values=brokers_list
    )
combo.place(x=50, y=50)
combo.set("Ativa")

button = ttk.Button(text="Selecionar", command=run)
button.place(x=50, y=100)

button2 = ttk.Button(text="Encerrar", command=encerrar)
button2.place(x=160, y=100)

main_window.mainloop()

if len(not_sent_emails) > 0:
    print(f"Não foi possível enviar email para os seguintes destinatários: {not_sent_emails}")
