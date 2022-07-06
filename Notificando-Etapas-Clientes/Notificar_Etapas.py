from pandas.core.base import DataError
from pandas.core.frame import DataFrame
from selenium import webdriver
import time
from selenium.webdriver.chrome import options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import PySimpleGUI as sg
import pandas as pd
from datetime import date
import xlsxwriter

def janela1():
    sg.theme('Reddit')
    layout = [
        [sg.Text('Comunicar clientes sobre o andamento das vistorias:',font=("Roboto", 17)),sg.Button('Duvidas',font=("Roboto", 11),key='Duvida1', button_color=("red", "white"))],
        [sg.Text('Conectar WhatsApp:',font=("Arial Bold", 14)),sg.Button('Conectar',font=("Roboto", 14),key="conectar")],
        [sg.Button('Enviar Menssagens',size=(50, 0),key='enviar',font=("Roboto", 20))]
    ]
    return sg.Window('Men Whats',layout=layout,finalize=True)

def janela2():
    sg.theme('Reddit')
    layout = [
        [sg.Text('Comunicar clientes sobre o andamento das vistorias:',font=("Roboto", 17)),sg.Button('Duvidas',font=("Roboto", 11),key='Duvida1', button_color=("red", "white"))],
        [sg.Text('Conectar WhatsApp:',font=("Arial Bold", 14)),sg.Button('Conectar',font=("Roboto", 14),key="conectar")],
        [sg.Button('Enviar Menssagens',size=(50, 0),key='enviar',font=("Roboto", 20))],
        [sg.Button('Enviar Todas Menssagens',size=(39, 0),key='enviar1',font=("Roboto", 20),button_color=("red", "white")),sg.Button('Duvidas',font=("Roboto", 20),size=(10, 0),key='Duvida2', button_color=("red", "white"))]
    ]
    return sg.Window('Men Whats',layout=layout,finalize=True)

# Definindo funções
def abrir_conversa(contato):
    campo_pesquisa = drive.find_element_by_xpath('//div[contains(@class,"copyable-text selectable-text")]')
    time.sleep(2)
    campo_pesquisa.click()
    campo_pesquisa.send_keys(contato)
    campo_pesquisa.send_keys(Keys.ENTER)
    time.sleep(1)

def enviar_mensagem(msg):
    time.sleep(1)
    campo_msg = drive.find_elements_by_xpath('//div[contains(@class,"copyable-text selectable-text")]')
    campo_msg[1].click()
    time.sleep(1)
    campo_msg[1].send_keys(msg)
    campo_msg[1].send_keys(Keys.ENTER)
    time.sleep(1)

def save(df_comp):
    df_resultado = df_comp
    data = '{}-{}-{}'.format(date.today().day, date.today().month, date.today().year)
    df_resultado.to_csv('Historico/Resultado do envio '+data+'.csv', encoding='utf-8', index=False)
    df_resultado.to_csv('Programa/Resultado.csv', encoding='utf-8', index=False)

def ler_tabelas():
    tabela1 = False
    tabela2 = False
    df = ''
    df_comp = ''
    try:
        df = pd.read_excel('H:/Meu Drive/Controle de Clientes.xlsx')
        tabela1 = True
    except:
        print('Erro ao ler tabela de controle')
        tabela1 = False     
    
    try:
        df_comp = pd.read_csv('Programa/Resultado.csv')
        # Acrescentando linhas na tabela caso faltem
        if len(df_comp.iloc[:,0]) != len(df.iloc[:,0]):
            for index in range(len(df_comp.iloc[:,0]),len(df.iloc[:,0])):
                df_comp.loc[index] = [df.iloc[index,0]] + [0,0] 
        # Organizando verificando todos os nomes da tabela
        for index in range(0,len(df_comp.iloc[:,0])):
             df_comp.loc[index,['Nome']] = df.iloc[index,0]
        tabela2 = True
        print('tabela de resultado lida/conferida')
    except:
        print('tabela de resultados nao encontrada criando uma nova')
        if tabela1:
            df_comp = DataFrame({'Nome':[],'Resultado':[],'Etapa':[]})
            for index in range(0,len(df.iloc[:,0])):
                df_comp.loc[index] = [df.iloc[index,0]] + [0,0]
            tabela2 = True
    if tabela1 and tabela2:
        tabelas_ok = True
    else:
        tabelas_ok = False                 

    return df,df_comp,tabelas_ok
    
def contagem_msg(df,df_comp,tabelas_ok):
    numero_envios = 0
    if tabelas_ok:
        for index in range(0,len(df)): 
            if df.iloc[index,5] != df_comp.iloc[index,2]:
                print(df.iloc[index,0],'-',df.iloc[index,5],'->',df_comp.iloc[index,2])
                numero_envios = numero_envios + 1
    return numero_envios

def enviar_menssagens(df,df_comp):
    numero_msg_enviadas = 0
    for index in range(0,len(df)): 
        try:
            if df.iloc[index,5] != df_comp.iloc[index,2]:
                telefone = int(df.iloc[index,4])
                abrir_conversa(telefone)
                if str(df.iloc[index,5]) == "Finalizado":
                    enviar_mensagem("Oi, passando para avisar que o seu processo na Enel foi  " + str(df.iloc[index,5]))
                else:
                    enviar_mensagem("Oi, passando para avisar que o seu processo na Enel foi atualizado, agora você está na etapa " + str(df.iloc[index,5]) + ", o prazo da Enel é ate o dia " + str(df.iloc[index,6]))
                print("\n")
                
                df_comp.loc[df_comp.index == index, 'Resultado'] = 'ATUALIZADA'
                df_comp.loc[df_comp.index == index, 'Etapa'] = df.iloc[index,5]
                print("Mensagem enviada para: ", str(df.iloc[index,0]), " - ", str(int(df.iloc[index,4])), " - " ,df.iloc[index,5]) 
                numero_msg_enviadas = numero_msg_enviadas + 1

        except:
            df_comp.loc[df_comp.index == index, 'Resultado'] = 'NAO ATUALIZADA'
            print("\n Menssagem não enviada para: " + str(df.iloc[index,0]))

    return df,df_comp,numero_msg_enviadas



# Inicializando variaveis ============================================================================================
df = ''
df_comp = ''
msg = ''
erro = 0
numero_envios = 0
conectado = 1
tabela = [0,0]

# Criando janela inicial =============================================================================================
janela1 = janela1()

while True:

    window,event,values = sg.read_all_windows()
    # Close ========================================================================================================================================
    if event == sg.WIN_CLOSED:
        break
    
    # Botão de Duvida =============================================================================================================================
    if event == 'Duvida1':
        sg.popup_ok('Este programa realiza o envio de mensagens automáticas para os clientes listados na tabela "Controle de Clientes" que deve estar localizada na mesma pasta, além disso o usuário deve possuir o chrome instalado.',font=("Roboto", 15))

    # Botão de Duvida =============================================================================================================================
    if event == 'Duvida2':
        sg.popup_ok("Foram 'enviadas'/'nao enviadas' as seguintes menssagens. \n",font=("Roboto", 15))

    # Botão de Conectar ============================================================================================================================

    if event == 'conectar':
        print("Conectando ao whatsapp")
        drive = webdriver.Chrome(ChromeDriverManager().install())
        drive.get('https://web.whatsapp.com/')
        conectado = 1

    # Botão de Enviar =============================================================================================================================

    if event == 'enviar':
        print('\n\n\n')
        print("Botão enviar pressionado")
        df=''
        df_comp=''
        df,df_comp,tabelas_ok = ler_tabelas()
        if not tabelas_ok:
            sg.popup_ok('Erro ao ler tabelas',font=("Roboto", 15))
        numero_envios = contagem_msg(df,df_comp,tabelas_ok)
                
        if sg.popup_ok_cancel('Serão Enviadas '+ str(numero_envios) + ' menssagens, continuar?',font=("Roboto", 12)) != 'Cancel':
            if tabelas_ok:
                if conectado == 1:
                    # Enviando menssagens
                    df,df_comp,numero_msg_enviadas = enviar_menssagens(df,df_comp)
                    save(df_comp)

                    sg.popup_ok('Numero de mensagens enviadas: '+str(numero_msg_enviadas),font=("Roboto", 15))
                    if numero_msg_enviadas != numero_envios:
                        sg.popup_ok('Erro ao enviar algumas menssagens: ' +str(numero_envios - numero_msg_enviadas),font=("Roboto", 15))
                else:
                    sg.popup_ok('Realize a conexão com o WhatsApp',font=("Roboto", 12))
            

            
            
       
