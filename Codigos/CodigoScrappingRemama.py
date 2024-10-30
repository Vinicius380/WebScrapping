'''

Instalar as bibliotecas:
    pip install webdriver-manager
    pip install selenium
    pip install pandas as pd
    pip install openpyxl
    
'''
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference

servico = Service()
options = webdriver.ChromeOptions()
navegador = webdriver.Chrome(service=servico, options=options)

url = 'https://www.instagram.com'

navegador.get(url)

login = "viniciusdejesussilva120@gmail.com"
senha = "zamasu120"

# Login no instagram
time.sleep(1)

navegador.find_element('xpath', '//*[@id="loginForm"]/div/div[1]/div/label/input').send_keys(login) # Inseri o E-mail no comboBox
time.sleep(2)
navegador.find_element('xpath', '//*[@id="loginForm"]/div/div[2]/div/label/input').send_keys(senha) # Inseri a Senha no ComboBox
time.sleep(2)
navegador.find_element('xpath', '//*[@id="loginForm"]/div/div[3]').click() # Clica no botão para efetuar o login
time.sleep(5)


navegador.find_element(By.CLASS_NAME, '_ac8f').click() # Clica no botão "Agora não" sobre Salvar suas informações de login
time.sleep(2)


navegador.get('https://www.instagram.com/remamadragaorosa/') # Entrando no instagram do REMAMA. 
time.sleep(2)

print("Status atual do Instagram do Remama:")

publicacoes = navegador.find_elements(By.CLASS_NAME,'_ac2a')[0].text # Pegando a quantidade de publicações
print(f'Quantidade de Publicações: {publicacoes}')
seguidores = navegador.find_elements(By.CLASS_NAME, '_ac2a')[1].text # Pegando a quantidade de seguidores
print(f'Numero de seguidores: {seguidores}')
seguindo = navegador.find_elements(By.CLASS_NAME, '_ac2a')[2].text # Pegando a quantidade de seguindo 
print(f'Quantidade que estão seguindo: {seguindo}')
time.sleep(2)

#Entra na postagem
link_postagem = navegador.find_elements(By.CLASS_NAME, "_aagu") # Pega a CLASS_NAME da primeira postagem
link_postagem[0].click()
time.sleep(2)

print("Dados da postagem")

# Capturando os dados 
titulo_postagem = navegador.find_elements(By.CLASS_NAME, "_a9zs")
titulo_postagem_text = titulo_postagem[0].text if titulo_postagem else "Título não encontrado"
#print(f'Titulo: {titulo_postagem_text}')

data_postagem = navegador.find_elements(By.CLASS_NAME, 'x1p4m5qa')
data_postagem_text = data_postagem[0].text if data_postagem else "Data não encontrada"
#print(f'Data: {data_postagem_text}')

visualizacoes = navegador.find_elements(By.CLASS_NAME, "_aauw")
visualizacoes_text = visualizacoes[0].text if visualizacoes else "Numero de visualizações não encontrado"
#print(f'Visualizações: {visualizacoes_text}')

dados = {
    'Publicações': [publicacoes],
    'Seguidores': [seguidores],
    'Seguindo': [seguindo],
    "Titulo": [titulo_postagem_text],
   "Visualizacoes": [visualizacoes_text],
    "Data": [data_postagem_text]
}
#Criacao Data Frame
df = pd.DataFrame(dados)

#Exportar Excel
arquivo_excel = "dados_instagram.xlsx"
df.to_excel(arquivo_excel, index=False)

#Carregar arquivo
wb = load_workbook(arquivo_excel)
ws = wb.active

#Criar grafico
chart = BarChart()
chart.title = "Estatísticas do Instagram"
chart.x_axis.title = "Métricas"
chart.y_axis.title = "Quantidade"

#Selecao de dados
data = Reference(ws, min_col=2, min_row=1, max_col=2, max_row=4)
categorias = Reference(ws, min_col=1, min_row=2, max_row=4)
chart.add_data(data, titles_from_data=True)
chart.set_categories(categorias)

#Adicionar a planilha
ws.add_chart(chart, "E5")

#Salvar arquivo
wb.save(arquivo_excel)

print("Dados e gráfico exportados para 'dados_instagram.xlsx'")