from selenium import webdriver
import time
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
import PySimpleGUI as sg
import pandas as pd

# Definindo funções
def abrir_conversa(contato):
    campo_pesquisa = drive.find_element_by_xpath('//div[contains(@class,"copyable-text selectable-text")]')
    time.sleep(1)
    campo_pesquisa.click()
    campo_pesquisa.send_keys(contato)
    campo_pesquisa.send_keys(Keys.ENTER)
    time.sleep(1)

def enviar_mensagem(msg):
    campo_msg = drive.find_elements_by_xpath('//div[contains(@class,"copyable-text selectable-text")]')
    campo_msg[1].click()
    time.sleep(1)
    campo_msg[1].send_keys(msg)
    campo_msg[1].send_keys(Keys.ENTER)
    time.sleep(1)

def close():
    campo_menu = drive.find_element_by_xpath('//div[contains(@title,"Menu")]')
    campo_menu.click()
    time.sleep(3)
    campo_log = drive.find_elements_by_xpath('//div[contains(@class,"_11srW _2xxet")]')
    campo_log[6].click()
    drive.close()


def janela_configuração():
    sg.theme('Reddit')
    layout = [
        [sg.Text('Texto da Mensagem (Arquivo de texto):',font=("Roboto", 21)),sg.Button('Importar Mensagem',font=("Roboto", 15))],
        [sg.Text('Planilha com telefones ou nomes dos contatos:',font=("Roboto", 17)),sg.Button('Importar',font=("Roboto", 15),key='Importar Numeros'),sg.Button('Atenção!', button_color=("red", "white"),font=("Roboto", 15),key='Duvidas')],
        [sg.Text('Caso os telefones não estejam na coluna 0 digite o numero da coluna:',font=("Arial Bold", 14)),sg.Input(key='coluna',size=(3,2),font=("Roboto", 14)),sg.Button('Confirmar',font=("Roboto", 14),key="Click Aqui")],
        [sg.Button('Enviar (Prepare o Celular Para Ler o QRCode do WhatsApp)',size=(50, 0),key='Enviar',font=("Roboto", 20))]
    ]
    return sg.Window('Men Whats',layout=layout,finalize=True)

# Criando janela
df = ''
msg = ''
number = 0
importe1 = False
importe2 = False
janela_configuração()

while True:
    window,event,values = sg.read_all_windows()
    # Close janelas
    if event == sg.WIN_CLOSED:
        break

    if event == 'Importar Mensagem':
        filename2 = sg.PopupGetFile('Get required file', no_window = True,file_types=(("text Files","*.txt"),))
        try:
            msg = open(filename2, "r")
            msg = msg.read() 
            importe1 = True
        except:
            sg.popup_error('Erro ao ler menssage',font=("Roboto", 15))

    if event == 'Importar Numeros':
        filename = sg.PopupGetFile('Get required file', no_window = True,file_types=(("CSV Files","*.csv"),))
        try:
            df = pd.read_csv(filename)
            importe2 = True
        except:
            sg.popup_error('Erro ao Ler Arquivo',font=("Roboto", 15))

    if event == 'Duvidas':
        sg.popup_ok('A planilha deve conter os numeros na primeira coluna, a partir da segunda celula. Pois a celula A0 é reservada para nomear a coluna.',font=("Roboto", 15))
    
    if event == 'Click Aqui':
        try:
            number = int(values['coluna'])
        except:
            number = values['coluna']   
        if number in range(0,1000):
            coluna_de_numeros = values['coluna']
            sg.popup_ok('Alteração Ralizada',font=("Roboto", 15))
        else:
            number = 0
            sg.popup_error('Digite Um Valor Valido',font=("Roboto", 15))
   
    if event == 'Enviar':
        if importe1 and importe2:
            telefones = df.iloc[:,number                                             ].tolist()
            print(telefones)
            # Acessando  o whats app
            drive = webdriver.Chrome(ChromeDriverManager().install())
            drive.get('https://web.whatsapp.com/')
            time.sleep(10)
            # Enviando menssagens
            for telefone in telefones:
                abrir_conversa(telefone)
                enviar_mensagem(msg)
            close()
            sg.popup_ok('Mensagens Enviadas',font=("Roboto", 15))
            window.close()
        else:
            sg.popup_error('Importe os Arquivos',font=("Roboto", 15))


