import requests
from bs4 import BeautifulSoup
import urllib3
import pandas as pd

# Suprimir avisos de solicitação insegura
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Solicitar a página ignorando a verificação de certificado
url = 'https://cesupinfra.intranet.bb.com.br/scap/#/painel-do-cliente/consulta-demandas/'
response = requests.get(url, verify=False)
soup = BeautifulSoup(response.content, 'html.parser')

# Imprimir o HTML da página
print(soup.prettify())

# Continue com o processamento dos dados
dados = []
for item in soup.find_all('div', class_='classe-exemplo'):
    titulo = item.find('h2').text if item.find('h2') else 'Sem Título'
    descricao = item.find('p').text if item.find('p') else 'Sem Descrição'
    dados.append({'Título': titulo, 'Descrição': descricao})

# Verificar se os dados foram extraídos corretamente
print(dados)

# Filtrar e armazenar os dados
df = pd.DataFrame(dados)
if 'Título' in df.columns:
    df_filtrado = df[df['Título'].str.contains('palavra-chave')]
    df_filtrado.to_csv('dados_filtrados.csv', index=False)
else:
    print("A coluna 'Título' não foi encontrada no DataFrame.")
