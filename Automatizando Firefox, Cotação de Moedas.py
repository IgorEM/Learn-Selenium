from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd

nav = webdriver.Firefox()

#pesquisar cotação dolar
nav.get("https://www.google.com/")

nav.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("Cotação Dólar Americano")
nav.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
time.sleep(2.5)
cotacao_dolar_eua = nav.find_element_by_xpath('/html/body/div[7]/div/div[9]/div[1]/div/div[2]/div[2]/div/div/div[1]/div/div/div/div/div[1]/div/div[1]/div[1]/div[2]/span[1]').get_attribute("data-value")
print("Cotação Dólar Americano=",cotacao_dolar_eua)

#abrir nova guia cotaçao euro
nav.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't') 
nav.get("https://www.google.com/") #CARREGAR NOVA ABA
nav.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("Cotação Euro")
nav.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
time.sleep(2.5)
cotacao_euro = nav.find_element_by_xpath('/html/body/div[7]/div/div[9]/div[1]/div/div[2]/div[2]/div/div/div[1]/div/div/div/div/div[1]/div/div[1]/div[1]/div[2]/span[1]').get_attribute("data-value")
print("Cotação Euro=",cotacao_euro)

#abrir nova guia cotaçao libra esterlina
nav.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't') 
nav.get("https://www.google.com/") #CARREGAR NOVA ABA
nav.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("Cotação Libra Esterlina")
nav.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
time.sleep(2.5)
cotacao_libra = nav.find_element_by_xpath('/html/body/div[7]/div/div[9]/div[1]/div/div[2]/div[2]/div/div/div[1]/div/div/div/div/div[1]/div/div[1]/div[1]/div[2]/span[1]').get_attribute("data-value")
print("Cotação Libra Esterlina=",cotacao_libra)

#abrir nova guia cotaçao Iene
nav.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't') 
nav.get("https://www.google.com/") #CARREGAR NOVA ABA
nav.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("Cotação Iene")
nav.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
time.sleep(2.5)
cotacao_iene = nav.find_element_by_xpath('/html/body/div[7]/div/div[9]/div[1]/div/div[2]/div[2]/div/div/div[1]/div/div/div/div/div[1]/div/div[1]/div[1]/div[2]/span[1]').get_attribute("data-value")
print("Cotação Iene=",cotacao_iene)

#abrir nova guia cotaçao OURO
nav.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't') 
nav.get("https://www.melhorcambio.com/") #CARREGAR NOVA ABA
aba_original = nav.window_handles[0] #permite "identificar" uma aba do navegador
#click #quando clica uma nova aba é aberta
nav.find_element_by_xpath('/html/body/div[14]/div[2]/div/table[2]/tbody/tr[2]/td[2]/a/img').click() #clicando na imagem

aba_nova = nav.window_handles[1]
nav.switch_to.window(aba_nova) #mudar para a nova aba
time.sleep(2.5)
cotacao_ouro = nav.find_element_by_id('comercial').get_attribute('value')
cotacao_ouro = cotacao_ouro.replace(",",".")
print("1g de ouro está custando",cotacao_ouro,"hoje")

nav.quit() #fecha navegador

#importar o excel
produtos_df = pd.read_excel(r'C:\Users\MeuUsuario\Desktop\Produtos.xlsx')
display(produtos_df)

#Encontradas as linhas que possuem “Dólar” na coluna “Moeda”, o Loc irá alterar a coluna  “Cotação”
produtos_df.loc[produtos_df['Moeda'] == "Dólar", "Cotação"] = float(cotacao_dolar_eua)

produtos_df.loc[produtos_df['Moeda'] == "Euro", "Cotação"] = float(cotacao_euro)
produtos_df.loc[produtos_df['Moeda'] == "Ouro", "Cotação"] = float(cotacao_ouro)

#recalculando os valores da tabela com as novas cotações
produtos_df['Preço Base Reais'] = produtos_df['Cotação'] * produtos_df['Preço Base Original']
produtos_df['Preço Final'] = produtos_df['Ajuste'] * produtos_df['Preço Base Reais']

display(produtos_df)

#exportar para um novo excel
produtos_df.to_excel(r'C:\Users\MeuUsuario\Desktop\Produtos Atualizados.xlsx', index=False) #Indica que não deverá ser exportado a coluna com o índice das linhas
