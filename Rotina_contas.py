#PARA SETAR OS PONTOS NA TELA, USE A FUNÇÃO: pyautogui.position()

import pandas as pd
import pyautogui
import clipboard
import pyperclip
import keyboard
import time

pyautogui.FAILSAFE = True

def copy_clipboard():
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(.01)  # ctrl-c is usually very fast but your program may execute faster
    return pyperclip.paste()


def conta_ativo():
    pyautogui.moveTo(433,636,duration=0.5)
    #pyautogui.click()
    #pyautogui.write('54244')############### Conta ativo
    #pyautogui.hotkey('enter')
    pyautogui.moveTo(218,290,duration=0.5) ##
    #pyautogui.click(button='right')        ##copia registro
    pyautogui.moveTo(243,425,duration=0.5)
    #pyautogui.click()
    pyautogui.moveTo(433,636,duration=0.5)#cancelar
    #pyautogui.click()
    pyautogui.moveTo(343,637,duration=0.5) #novo
    #pyautogui.click()
    pyautogui.moveTo(218,290,duration=0.5) ##
    #pyautogui.click(button='right')
    pyautogui.moveTo(271,441,duration=0.5) ##colar registro
    #pyautogui.click()
    pyautogui.moveTo(206,255,duration=0.5) ##descrição
    #pyautogui.click()

def conta_passivo():
    pyautogui.moveTo(433,636,duration=0.5)
    #pyautogui.click()
    #pyautogui.write('54245')############### Conta passivo
    #pyautogui.hotkey('enter')
    pyautogui.moveTo(218,290,duration=0.5) ##
    #pyautogui.click(button='right')        ##copia registro
    pyautogui.moveTo(243,425,duration=0.5)
    #pyautogui.click()
    pyautogui.moveTo(433,636,duration=0.5)#cancelar
    #pyautogui.click()
    pyautogui.moveTo(343,637,duration=0.5) #novo
    #pyautogui.click()
    pyautogui.moveTo(218,290,duration=0.5) ##
    #pyautogui.click(button='right')
    pyautogui.moveTo(271,441,duration=0.5) ##colar registro
    #pyautogui.click()
    pyautogui.moveTo(206,255,duration=0.5) ##descrição
    #pyautogui.click()
  



def inicio():
    print("Bem vindo ao processo de robotização para criação das contas dos processo de Importação")
    print("Escreva de qual linha do Excel deseja iniciar o processo de importação")
    entrada = input()
    entrada=(int(entrada))
    entrada= entrada-2
    if entrada >=0 and entrada <= len(num_DI):
        if len(num_DI)== len(ref_DI):
            n=entrada
            #for n in range(len(num_DI)):
            for i in range(n, len(ref_DI), 1):
                descricao_adto = 'Adto de Importacao Filial - DI'+' ' + num_DI[i].strip()+' ' + ref_DI[i].strip()
                descricao_import = 'IMPORTACAO FILIAL - DI'+' ' + num_DI[i].strip()+' ' + ref_DI[i].strip()
                conta_ativo()
                pyautogui.write(descricao_adto)
                print(descricao_adto)
                conta_passivo()
                pyautogui.write(descricao_import)
                print(descricao_import)
                #pyautogui.write(descricao)
            #   num_DI_pos = (num_DI[i])
            #  print(num_DI_pos)
            # ref_DI_pos = (ref_DI[i])
                #print(ref_DI_pos)
        else:
            print("O numero de linhas do Número DI é diferente do Número de Linhas da Referência")

    else:
        print("Você digitou um número de linhas incompatível")



x = pd.read_excel("test.xlsx")
num_DI = x['Número DI']
ref_DI = x['Referência']
#print(num_DI)
#print(len(num_DI))
#print(len(ref_DI))
inicio()


