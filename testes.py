import pandas as pd
import json
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo

# Caminho do arquivo JSON
json_file_path = "testemes.json"

try:
    # Carregar o JSON a partir do arquivo
    with open(json_file_path, 'r', encoding='utf-8') as file:
        data = json.load(file)

    # Converter os dados JSON em um DataFrame do pandas
    df = pd.DataFrame(data)

    # Remover as colunas 'email', 'Nome' e 'quantidade_participantes'
    df.drop(columns=['email', 'Nome', 'quantidade_participantes'], inplace=True)

    # Criar um novo Workbook e selecionar a planilha ativa
    wb = Workbook()
    ws = wb.active

    # Adicionar os dados do DataFrame à planilha
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    # Definir o intervalo da tabela (todas as células com dados)
    tab = Table(displayName="TabelaEventos", ref=f"A1:{chr(65+len(df.columns)-1)}{len(df)+1}")

    # Adicionar estilo à tabela
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style

    # Adicionar a tabela à planilha
    ws.add_table(tab)

    # Salvar o Workbook como um arquivo Excel
    excel_path = "eventos.xlsx"
    wb.save(excel_path)

    print(f"O arquivo Excel foi salvo como {excel_path}.")
except json.JSONDecodeError as e:
    print(f"Erro ao decodificar o JSON: {e}")
except FileNotFoundError:
    print(f"Arquivo não encontrado: {json_file_path}")
except Exception as e:
    print(f"Ocorreu um erro: {e}")