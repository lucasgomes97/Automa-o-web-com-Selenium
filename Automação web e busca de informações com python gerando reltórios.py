import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd


navegador = selenium.webdriver.Chrome()
# 1 pegar as cotações do dolar, euro e ouro
# Euro
navegador.get("https://www.google.com.br/")
navegador.find_element('xpath',
                       '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("Cotação euro")
navegador.find_element('xpath',
                       '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cotacao_euro = navegador.find_element('xpath',
                       '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(cotacao_euro)

# Dolar
navegador.get("https://www.google.com.br/")
navegador.find_element('xpath',
                       '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("Cotação dolar")
navegador.find_element('xpath',
                       '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cotacao_dolar = navegador.find_element('xpath',
                       '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(cotacao_dolar)

# Ouro
navegador.get("https://www.melhorcambio.com/ouro-hoje")
cotacao_ouro = navegador.find_element('xpath', ('//*[@id="comercial"]')).get_attribute("value")
cotacao_ouro = cotacao_ouro.replace(",", ".")
print(cotacao_ouro)
navegador.quit()

# 2 Importar e Atualizar a base de dados
tabela = pd.read_excel(r'C:\Users\vidal\Downloads\Produtos.xlsx')
print(tabela)
# 2.1atualizando os valorer
tabela.loc[tabela["Moeda"] == "Dólar", "Cotação"] = float(cotacao_dolar)
tabela.loc[tabela["Moeda"] == "Euro", "Cotação"] = float(cotacao_euro)
tabela.loc[tabela["Moeda"] == "Ouro", "Cotação"] = float(cotacao_ouro)

# 3 Recalcular os preços ( preço de venda  = preço de compra * margem) ,(preço de compra = cotação * preço original)
tabela["Preço de Compra"] = tabela["Cotação"] * tabela["Preço Original"]
tabela["Preço de Venda"] = tabela["Preço de Compra"] * tabela["Margem"]

# 4 Exportar a base atualizada
tabela.to_excel(r'C:\Users\vidal\Downloads\Produtos.xlsx', index=False)
