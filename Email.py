import PySimpleGUI as sg
import pandas as pd
import smtplib
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import config

def email(destinatario,titulo,msg,user,password):
    host = 'smtp.gmail.com'
    port = 587

    # Criando objeto
    server = smtplib.SMTP(host, port)

    # Login com servidor
    server.ehlo()
    server.starttls()
    server.login(user, password)

    # Criando mensagem
    message = msg
    email_msg = MIMEMultipart()
    email_msg['From'] = user
    email_msg['To'] = destinatario
    email_msg['Subject'] = titulo
    email_msg.attach(MIMEText(message, 'plain'))

    server.sendmail(email_msg['From'], email_msg['To'], email_msg.as_string())
    print('Mensagem enviada! - ',destinatario)
    server.quit()

coluna_emails = 0 
lista_emails = []
Importe = False

def janela_login():
    sg.theme('Reddit')
    layout = [
        [sg.Text('Email:'),sg.Input(key='user')],
        [sg.Text('Senha:'),sg.Input(key='senha',password_char='*')],
        [sg.Button('Entrar')]
    ]
    return sg.Window('login',layout=layout,size=(290, 100),finalize=True)

def janela_email():
    sg.theme('Reddit')
    layout = [
        [sg.Text('Titulo:'),sg.Input(key='titulo')],
        [sg.Text('Texto do Email:'),sg.Button('Importar Mensagem')],
        [sg.Text('CSV - Emails:'),sg.Button('Importar Arquivo com Emails')],
        [sg.Button('Voltar'),sg.Button('Enviar')]
    ]
    return sg.Window('Moentar Pedido',layout=layout,size=(290, 100),finalize=True)

# Criando Janelas, e habilitando janela1
janela1, janela2 = janela_login(), None

# Loop do programa
while True:
    window,event,values = sg.read_all_windows()

    # Close janelas
    if window == janela1  and event == sg.WIN_CLOSED:
        break
    if window == janela2  and event == sg.WIN_CLOSED:
        break
    
    #Testando login
    if window == janela1 and event == 'Entrar': # Nome do botão
        if values['user'] != '' and values['senha'] != '':
            janela2 = janela_email()
            janela1.hide()
            user = values['user']
            senha = values['senha']
        else :
            sg.popup('Digite um Email e Senha')
    #Voltando da janela2
    if window == janela2 and event == 'Voltar':
        janela2.hide()
        janela1.un_hide()
    
    # Enviando Email
    if window == janela2 and event == 'Enviar':
        print('Enviar')
        if values['csv'] != "":
            print('Digitado')
        else:
            print('Não digitou nada')

    if window == janela2 and event == 'Importar Mensagem':
        filename2 = sg.PopupGetFile('Get required file', no_window = True,file_types=(("text Files","*.txt"),))
        try:
            msg = open(filename2, "r")
            msg = msg.read() 
            sg.popup('Reading menssage')
        except:
            sg.popup_error('Error reading menssage')

    # Enviando Importar arquivo
    if window == janela2 and event == 'Importar Arquivo com Emails':
        filename = sg.PopupGetFile('Get required file', no_window = True,file_types=(("CSV Files","*.csv"),))
        try:
            df = pd.read_csv(filename)
            Importe = True
            window.close()
            break
        except:
            sg.popup_error('Error reading CSV')
        
if Importe :
    emails = df.iloc[:,coluna_emails].tolist()
    # Criando lista de amails unicos
    for i in emails:
        email(i,values['titulo'],msg,user,senha);
    print('Emails Enviados Para:',emails)


