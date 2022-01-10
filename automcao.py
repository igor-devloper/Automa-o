import pandas as pd
import pyautogui as pi
import time
from datetime import date, datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pyperclip as i




pi.PAUSE = 1
pi.alert("Vai começar, aperte OK e não mexa em nada")


nav = webdriver.Chrome()
nav.get('https://www.google.com.br/')
nav.find_element_by_xpath('//*[@id="gb"]/div/div[2]/a').click()
email = "cleoberto.wagner@energiaarion.com.br"
i.copy(email)
pi.hotkey("ctrl", "v")
pi.press("enter")
time.sleep(4)
senha = "cws231078"
i.copy(senha)
pi.hotkey("ctrl", "v")
pi.press("enter")
time.sleep(5)
pi.hotkey('ctrl', 't')
pi.click(371, 71)
link = "https://trello.com/b/lDMv1rIH/adequa%C3%A7%C3%A3o"
i.copy(link)
pi.hotkey("ctrl", "v")
pi.press("enter")
time.sleep(7)
pi.click(983, 162)
time.sleep(5)
pi.click(510, 574)
time.sleep(60)
pi.click(600, 229)
time.sleep(40)
pi.click(275,324)
pi.click(276, 414)
pi.click(1033, 679, clicks=11)
time.sleep(5)
pi.click(906, 645)
time.sleep(3)
pi.click(391, 490)
pi.click(723, 542 )
pi.click(549, 528)
pi.write("Card URL")
pi.press('enter')
pi.write("Card Name")
pi.press('enter')
pi.write("List Name")
pi.press('enter')
pi.write("Labels")
pi.press('enter')
pi.write("Due Date")
pi.press('enter')
pi.write("Location")
pi.press('enter')
pi.write("Last Activity Date")
pi.press('enter')
pi.write("Custom Fields")
pi.press('enter')
pi.click(607, 589)
time.sleep(20)
pi.click(182, 684)
time.sleep(50)
pi.click(1154, 84)
time.sleep(5)
pi.click(594, 288, button='right')
time.sleep(3)
pi.click(669, 445)
pi.click(638, 493)
time.sleep(3)
pi.hotkey('ctrl', 'b')
pi.click(771, 755)
time.sleep(10)
#abrir o excel e excluir o adequação



           








#ler base de dados
tabelaInicial = pd.read_excel("ADEQUAÇÃO 20211023")
#tratar base de dados
tabelaInicial = tabelaInicial.drop("Due Date Status", axis=1)
tabelaInicial = tabelaInicial.drop("Custom Field: CONTRATO DU", axis=1)
tabelaInicial = tabelaInicial.drop("Custom Field: CONTRATO PROJ SMF", axis=1)
tabelaInicial = tabelaInicial.drop("Custom Field: CONTRATO ADQ SMF", axis=1)
tabelaInicial = tabelaInicial.drop("Custom Field: FOR. KIT PADRÃO", axis=1)
tabelaInicial = tabelaInicial.drop("Custom Field: CLIENTE", axis=1)
tabelaInicial = tabelaInicial.drop("Custom Field: Nº Loja / Unidade", axis=1)
tabelaInicial = tabelaInicial.drop("Custom Field: CNPJ", axis=1)
tabelaInicial = tabelaInicial.drop("Custom Field: UC", axis=1)
tabelaInicial = tabelaInicial.drop("Custom Field: TIPO DE LOJA", axis=1)
tabelaInicial = tabelaInicial.drop("Custom Field: CONTATO", axis=1)
tabelaInicial = tabelaInicial.drop("Custom Field: TIPO DE FATURAMENTO", axis=1)
tabelaInicial = tabelaInicial.drop("Custom Field: RM DE INFRA", axis=1)
tabelaInicial = tabelaInicial.drop("Custom Field: RM DE PAINEL", axis=1)
tabelaInicial = tabelaInicial.drop("Custom Field: RM DE COMPONENTES", axis=1)
tabelaInicial = tabelaInicial.drop("Custom Field: NF FATURAMENTO", axis=1)
tabelaInicial = tabelaInicial.drop("Custom Field: DATA DO D.U", axis=1)
tabelaInicial = tabelaInicial.drop("Custom Field: TIPO DE MEDIÇÃO", axis=1)
tabelaInicial.to_excel("tabela atualizada.xlsx", index=False)


print(tabelaInicial.info())
display(tabelaInicial)


pi.PAUSE = 1



#copiar base de dados
pi.click(63, 745)
rt = "tabela atualizada.xlsx"
i.copy(rt)
pi.hotkey('ctrl', 'v') 
pi.press('enter', presses=2)
time.sleep(12)
pi.press('Down')
pi.keyDown('shift')  
pi.press('right', presses=4)
pi.keyUp('shift')
pi.hotkey('ctrl', 'shift', 'Down')
pi.hotkey('ctrl', 'c')
pi.hotkey('ctrl', 'home')
time.sleep(10)
pi.click(x=754, y=753)
pi.hotkey('ctrl', 't')
pi.write(r"https://docs.google.com/spreadsheets/d/1XYDJFC6kuoJ3zlEWQtUCqnVMpGsef1eyMchF-UV7zWI/edit#gid=1783529448")
pi.press('enter')
time.sleep(30)
pi.press('Down')
pi.keyDown('shift')
pi.press('right')
pi.press('right')
pi.press('right')
pi.press('right')
pi.keyUp('shift')
pi.hotkey('ctrl', 'shift', 'Down')
pi.press('del')

pi.hotkey('ctrl', 'home')
pi.press('Down')
pi.hotkey('ctrl', 'shift', 'v')
time.sleep(10)



#copiar celula por celula




nav.quit()

#1 é a location
pi.click(x=894, y=754)
pi.click(x=994, y=660)
pi.click(x=636, y=256)
pi.hotkey('ctrl', 'c')
pi.click(x=754, y=751)
pi.press('right')
pi.press('right')
pi.press('right')
pi.press('right')
pi.press('right')
pi.click(x=750, y=210)
pi.hotkey('ctrl', 'shift', 'v')
time.sleep(5)
pi.click(x=273, y=228)
pi.click(x=126, y=154)
pi.click(x=607, y=229)

#2 é a last activit date
pi.press('right')
pi.click(x=894, y=754)
pi.click(x=994, y=660)
pi.click(x=707, y=255)
pi.hotkey('ctrl', 'c')
pi.click(x=754, y=751)
pi.click(x=424, y=215)
pi.hotkey('ctrl', 'shift', 'v')


#3 é a data de data de exexução
pi.press('right')
pi.click(x=894, y=754)
pi.click(x=994, y=660)
pi.click(x=572, y=258)
pi.hotkey('ctrl', 'c')
pi.click(x=754, y=751)
pi.click(x=424, y=215)
pi.hotkey('ctrl', 'shift', 'v')

#4 é a data de migração 
pi.press('right')
pi.click(x=894, y=754)
pi.click(x=994, y=660)
pi.click(x=436, y=255)
pi.hotkey('ctrl', 'c')
pi.click(x=754, y=751)
pi.click(x=424, y=215)
pi.hotkey('ctrl', 'shift', 'v')
pi.press('right')
pi.click(x=1007, y=680)
pi.click(x=1007, y=680)
#5 é a concessionaria
pi.press('right')
pi.click(x=894, y=754)
pi.click(x=994, y=660)
pi.click(x=382, y=258)
pi.hotkey('ctrl', 'c')
pi.click(x=754, y=751)
pi.click(x=424, y=215)
pi.hotkey('ctrl', 'shift', 'v')
#6 é a codigo GEOB 
pi.press('right')
pi.click(x=894, y=754)
pi.click(x=994, y=660)
pi.click(x=513, y=259)
pi.hotkey('ctrl', 'c')
pi.click(x=754, y=751)
pi.click(x=838, y=213)
pi.hotkey('ctrl', 'shift', 'v')
pi.hotkey('ctrl','home')

time.sleep(10)

pi.hotkey('ctrl', 'f')
time.sleep(3)
#clicar nos 3 pontos
pi.click(1199, 210)


#1
pi.write(r" de janeiro de ")
pi.click(698, 306)
pi.write(r"/01/")
pi.click(561, 368)
pi.press('Down')
pi.press('enter')
pi.click(746, 603)


#2
time.sleep(2)
pi.click(687, 248, clicks=3)
pi.write(r" de fevereiro de ")
pi.click(698, 306, clicks=3)
pi.write(r"/02/")
pi.click(746, 603)


#3
time.sleep(2)
pi.click(687, 248, clicks=3)
mes = " de março de "
i.copy(mes)
pi.hotkey("ctrl", "v")
pi.press("enter")
pi.click(698, 306, clicks=3)
pi.write(r"/03/")
pi.click(746, 603)



#1
time.sleep(2)
pi.click(687, 248, clicks=3)
pi.write(r" de abril de ")
pi.click(698, 306, clicks=3)
pi.write(r"/04/")
pi.click(746, 603)



#2
time.sleep(2)
pi.click(687, 248, clicks=3)
pi.write(r" de maio de ")
pi.click(698, 306, clicks=3)
pi.write(r"/05/")
pi.click(746, 603)


#3
time.sleep(2)
pi.click(687, 248, clicks=3)
pi.write(r" de junho de ")
pi.click(698, 306, clicks=3)
pi.write(r"/06/")
pi.click(746, 603)



#1
time.sleep(2)
pi.click(687, 248, clicks=3)
pi.write(r" de julho de ")
pi.click(698, 306, clicks=3)
pi.write(r"/07/")
pi.click(746, 603)


#2
time.sleep(2)
pi.click(687, 248, clicks=3)
pi.write(r" de agosto de ")
pi.click(698, 306, clicks=3)
pi.write(r"/08/")
pi.click(746, 603)


#3
time.sleep(2)
pi.click(687, 248, clicks=3)
pi.write(r" de setembro de ")
pi.click(698, 306, clicks=3)
pi.write(r"/09/")
pi.click(746, 603)



#1
time.sleep(2)
pi.click(687, 248, clicks=3)
pi.write(r" de outubro de ")
pi.click(698, 306, clicks=3)
pi.write(r"/10/")
pi.click(746, 603)


#2
time.sleep(2)
pi.click(687, 248, clicks=3)
pi.write(r" de novembro de ")
pi.click(698, 306, clicks=3)
pi.write(r"/11/")
pi.click(746, 603)


#3
time.sleep(2)
pi.click(687, 248, clicks=3)
pi.write(r" de dezembro de ")
pi.click(698, 306, clicks=3)
pi.write(r"/12/")
pi.click(746, 603)

pi.click(x=884, y=606)

time.sleep(5)

pi.hotkey('ctrl', 'home')
pi.press('right')
pi.press('right')
pi.press('right')
pi.press('right')
pi.click(x=736, y=211)
pi.click(x=352, y=154)
pi.click(x=467, y=528)
pi.press('right')
pi.press('right')
pi.click(x=392, y=214)
pi.click(x=352, y=154)
pi.click(x=467, y=528)
pi.press('right')
pi.click(x=392, y=214)
pi.click(x=352, y=154)
pi.click(x=467, y=528)
pi.press('right')
pi.click(x=392, y=214)
pi.click(x=352, y=154)
pi.click(x=467, y=528)

time.sleep(3)
pi.hotkey('ctrl', 'home')
pi.press('right')
pi.press('right')
pi.press('right')
pi.press('right')
pi.press('right')
pi.press('right')
pi.press('right')
pi.click(x=487, y=235, clicks=2)
pi.hotkey('ctrl', 'a')
DT = "DATA EXECUÇÃO"
i.copy(DT)
pi.hotkey("ctrl", "v")
pi.press('enter')

pi.press('up')
pi.press('right')
pi.click(x=487, y=235, clicks=2)
pi.hotkey('ctrl', 'a')
DM = "DATA/MIGRAÇÃO"
i.copy(DM)
pi.hotkey("ctrl", "v")
pi.press('enter')
pi.press('up')
pi.press('right')
pi.click(x=487, y=235, clicks=2)
pi.hotkey('ctrl', 'a')
CS = "Concessionária"
i.copy(CS)
pi.hotkey("ctrl", "v")
pi.press('enter')
pi.press('up')
pi.press('right')
pi.click(x=912, y=227, clicks=2)
pi.hotkey('ctrl', 'a')
CG = "Código GEOB"
i.copy(CG)
pi.hotkey("ctrl", "v")
pi.press('enter')
pi.press('right')


pi.alert("o codigo acabou, o computador voltou a ser seu :)!!!")




