# importa biblioteca
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import pyscreeze

webbrowser.open('https://web.whatsapp.com')
sleep(15)

dados = openpyxl.load_workbook('Assets\Dados.xlsx')
paginas_aniversariantes = dados['dados']

def new_func():
    seta = pyscreeze.locateCenterOnScreen('botao.png')
    return seta

for linha in paginas_aniversariantes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    nascimento = linha[2].value
    mensagem = f'Olá {nome} parabéns'   
    
    try:
        link_mensagem_whatsaap = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'    
        webbrowser.open(link_mensagem_whatsaap)
        sleep(10)
        seta = new_func()
        sleep(5)
        #seta = pyautogui.locateCenterOnScreen('botao.png')
        #sleep(5)
        pyautogui.click(seta[0],seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl','W')
        sleep(5)
    except: 
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv', 'a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome}, {telefone}') 
        sleep(5)     
        pyautogui.hotkey('ctrl','W')
        sleep(5)     
