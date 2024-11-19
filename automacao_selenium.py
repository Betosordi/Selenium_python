from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
import pyperclip
from datetime import datetime, timedelta
import pandas as pd
import os

# Parte para obter a data ***************************************************************************
# Obter a data e hora atuais
data_atual = datetime.now()

# Obter o ano atual
ano_atual = datetime.now().year

# Obter o mês atual
mes_atual = data_atual.month

# Obter a data e hora atuais
data_atual = datetime.now()

# Subtrair um dia da data atual (dia de ontem)
ontem = data_atual - timedelta(days=1)

# Obter a data de ontem no formato 'dd/mm/yyyy'
data_ontem_com_barras = ontem.strftime("%d/%m/%Y")

# Substituir o dia pela data 1 para obter o primeiro dia do mês
primeiro_dia_mes = data_atual.replace(day=1)

# Formatar a data no formato 'dd/mm/yyyy'
primeiro_dia_mes_atual = primeiro_dia_mes.strftime("%d/%m/%Y")

#*******************************************************************************

# Caminho onde você extraiu o ChromeDriver
caminho_webdriver = 'C:/Users/Assistente TI/chromedriver-win64/chromedriver-win64/chromedriver.exe'

# Crie um objeto Service com o caminho do ChromeDriver
service = Service(executable_path=caminho_webdriver)

# Inicializa o navegador Chrome com o Service
driver_iniciar = webdriver.Chrome(service=service)

# Abre a página de login
driver_iniciar.get("link da página da web")
######################################################################################################
time.sleep(5)

# Simulando comandos na pagina
# Campo de usuario
campo_usuario = driver_iniciar.find_element(By.XPATH, "/html/body/app-root/div[3]/app-tela-login/div/form/div/div[1]/app-input/div/div/div[2]/input")
campo_usuario.send_keys("seu usuario")
time.sleep(5)
#-------------------------------------------------------------------------------------------------
# Campo de senha
campo_senha = driver_iniciar.find_element(By.XPATH, "/html/body/app-root/div[3]/app-tela-login/div/form/div/div[2]/app-input/div/div/div[2]/input")
campo_senha.send_keys("sua senha")
time.sleep(5)
#-------------------------------------------------------------------------------------------------
# Localizando o botão de login usando o XPath e clicando
botao_login = driver_iniciar.find_element(By.XPATH, "/html/body/app-root/div[3]/app-tela-login/div/form/div/div[5]/app-button/div/button").click()
time.sleep(8)
#-----------------------------------------------------------------------
# Clicando na aba RELATÓRIO
relatorio = driver_iniciar.find_element(By.XPATH, '//*[@id="Relatorios"]/strong').click()
time.sleep(5)
#-----------------------------------------------------------------------
# Clicando no campo vendas
campo_vendas = driver_iniciar.find_element(By.XPATH, '//*[@id="ConteudoHiper"]/div[1]/div/div/div/div/div/div/div/div[2]/ul/li[1]').click()
time.sleep(5)
#-----------------------------------------------------------------------

# Quantidade de pixels para rolar para baixo
scroll_amount = 300

# Executa o scroll para baixo por 300 pixels
driver_iniciar.execute_script(f"window.scrollBy(0, {scroll_amount});")
time.sleep(5)
#---------------------------------------------------------

# Clicando no campo produtos
campo_produtos_vendidos = driver_iniciar.find_element(By.XPATH, '(//*[@id="aRelatorio"])[15]').click()
time.sleep(5)
#------------------------------------------------------------------------------

# Clicando no campo flegue Este dia
campo_flegue = driver_iniciar.find_element(By.XPATH, '/html/body/div[2]/div[1]/div/div/div/div/div[2]/div/div/div[1]/div/div/div[2]/div/div[2]/div/label').click()
time.sleep(5)
#------------------------------------------------------------------------------

# Apagando a data
data_apagando = driver_iniciar.find_element(By.XPATH, '/html/body/div[2]/div[1]/div/div/div/div/div[2]/div/div/div[1]/div/div/div[2]/div/div[1]/div[1]/span[2]/i').click()
time.sleep(5)
#------------------------------------------------------------------------------

# Colocando a data inicial
campo_data = driver_iniciar.find_element(By.XPATH, '/html/body/div[2]/div[1]/div/div/div/div/div[2]/div/div/div[1]/div/div/div[2]/div/div[1]/div[1]/input[1]').send_keys(primeiro_dia_mes_atual)  # Substitua pelo texto que você deseja inserir
time.sleep(5)
#------------------------------------------------------------------------------

# Colocando a data final
campo_data_final = driver_iniciar.find_element(By.XPATH, '/html/body/div[2]/div[1]/div/div/div/div/div[2]/div/div/div[1]/div/div/div[3]/div/div[1]/input[1]').send_keys(data_ontem_com_barras)  # Substitua pelo texto que você deseja inserir
time.sleep(5)
#------------------------------------------------------------------------------
# Clicando em aplicar
campo_aplicar = driver_iniciar.find_element(By.XPATH, '/html/body/div[2]/div[1]/div/div/div/div/div[2]/div/div/div[3]/button[2]').click()
time.sleep(60)
#------------------------------------------------------------------------------
# Parte do JavaScript integrando com o Python, dando ctrl c na Tag que possui a lista das vendas de produtos
# Obter o conteúdo do elemento
content = driver_iniciar.execute_script("return document.querySelector('.table-responsive').innerText;")

# Copiar o conteúdo para a área de transferência
pyperclip.copy(content)

time.sleep(10)
#***********************************************************************************************************************************************

# Criar um dicionário que mapeia números de mês para o nome do mês
meses = {
    1: "janeiro",
    2: "fevereiro",
    3: "março",
    4: "abril",
    5: "maio",
    6: "junho",
    7: "julho",
    8: "agosto",
    9: "setembro",
    10: "outubro",
    11: "novembro",
    12: "dezembro"
}

# Obter o nome do mês atual usando o número
mes_atual_nome = meses[mes_atual]

# Transformando o que copiei da tag table-responsive em um arquivo csv.
# Definir o caminho onde o arquivo será salvo
caminho_arquivo = f"C:/Users/seu usuario/Meu Drive/Dados para BI/vendas_{mes_atual_nome}_{ano_atual}.csv"  # Substitua pelo caminho desejado

# Converter o conteúdo para formato CSV (opcional, dependendo do formato do conteúdo)
# Assumindo que as colunas são separadas por tabulação e as linhas por quebras de linha
linhas = content.split("\n")  # Divide o texto em linhas
dados_csv = [linha.split("\t") for linha in linhas]  # Divide cada linha em colunas por tabulação

# Salvar o conteúdo no arquivo CSV
with open(caminho_arquivo, "w", encoding="utf-8", newline="") as arquivo_csv:
    for linha in dados_csv:
        arquivo_csv.write(";".join(linha) + "\n")  # Escreve os dados no formato CSV

###########################################################################################################################################

# Acrescenta a coluna Mês no dataframe
caminho_dos_meses = f"C:/Users/seu usuario/Meu Drive/Dados para BI/vendas_{mes_atual_nome}_{ano_atual}.csv"

mes = pd.read_csv(caminho_dos_meses,sep = ';')

# Adicionar a coluna 'Mês' com o valor 'novembro' para todas as linhas
mes['Mês'] = mes_atual_nome

# Salvando o dataframe
caminho_salvar = f"C:/Users/seu usuario/Meu Drive/Dados para BI/vendas_{mes_atual_nome}_{ano_atual}.csv"

mes.to_csv(caminho_salvar, sep = ';')
######################################################################################################################################

# Unificando os arquivos dos meses

# Caminho onde estão os arquivos
caminho_dos_arquivos = f"C://Users/seu usuario/Meu Drive/Dados para BI"

# Criar uma lista para armazenar os DataFrames
lista_dataframes = []

# Percorrer os arquivos na pasta
for arquivo in os.listdir(caminho_dos_arquivos):
    # Verificar se o arquivo começa com 'vendas' e termina com '.csv'
    if arquivo.startswith("vendas") and arquivo.endswith(".csv"):
        caminho_completo = os.path.join(caminho_dos_arquivos, arquivo)
        
        # Ler o arquivo CSV e adicionar à lista
        df = pd.read_csv(caminho_completo, sep=';')  # Ajuste o separador conforme necessário
        lista_dataframes.append(df)

# Unificar todos os DataFrames em um único
df_unificado = pd.concat(lista_dataframes, ignore_index=True)

caminho_para_salvar = f"C:/Users/seu usuario/Meu Drive/Dados para BI/Dados unificados/vendas_franquiados.csv"

df_unificado['Número de vendas'] = df_unificado['Número de vendas'].astype(str)
df_unificado['Quantidade vendida'] = df_unificado['Quantidade vendida'].astype(str)

# Remover '.0' do final de cada valor
df_unificado['Número de vendas'] = df_unificado['Número de vendas'].str.rstrip('.0')

# Excluir a coluna 'Unnamed: 0'
df_unificado = df_unificado.drop(columns=['Unnamed: 0'])

# Excluir a coluna 'Unnamed: 0'
df_unificado = df_unificado.drop(columns=['#'])

# Excluir a última linha usando o índice
df_unificado = df_unificado.drop(df_unificado.index[-1])

# Salvar o DataFrame unificado em um novo arquivo CSV
caminho_saida = os.path.join(caminho_para_salvar)
df_unificado.to_csv(caminho_saida, sep=';', index=False, encoding='utf-8')

#######################################################################################################################################################
