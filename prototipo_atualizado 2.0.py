#====================ENCAPSULAMENTO DE VERIFICAÇÃO DE BIBLIOTECA======================#

import sys
import subprocess

# Função para verificar se um pacote está instalado e, se não estiver, retorna False
def verificar_e_instalar(pacote):
    try:
        __import__(pacote)  # Tenta importar o pacote
    except ImportError:
        return False  # Se a importação falhar, retorna False
    return True  # Se a importação for bem-sucedida, retorna True

# Função para verificar todas as bibliotecas necessárias
def verificar_bibliotecas():
    # Lista de pacotes necessários
    pacotes_necessarios = ["numpy", "pandas", "requests", "matplotlib","openpyxl"]
    pacotes_faltando = []  # Lista para armazenar pacotes que não estão instalados

    # Verifica cada pacote na lista de pacotes necessários
    for pacote in pacotes_necessarios:
        if not verificar_e_instalar(pacote):  # Se o pacote não estiver instalado
            pacotes_faltando.append(pacote)  # Adiciona o pacote à lista de pacotes faltando

    # Se não houver pacotes faltando, retorna e não exibe a mensagem
    if not pacotes_faltando:
        return

    # Exibe a mensagem informando que as bibliotecas necessárias não estão instaladas
    print('\033[1m Parece que você não tem as Bibliotecas necessárias para rodar o nosso programa. \033[0m\n')
    # Solicita ao usuário se deseja instalar as bibliotecas
    escolha_biblioteca = input('Deseja instalar as Bibliotecas? (Y/N): ')

    # Se o usuário escolher instalar as bibliotecas
    if escolha_biblioteca == 'y' or escolha_biblioteca == 'Y':
        for pacote in pacotes_faltando:  # Para cada pacote faltando
            subprocess.check_call([sys.executable, "-m", "pip", "install", pacote])  # Instala o pacote usando pip
    # Se o usuário escolher não instalar as bibliotecas
    elif escolha_biblioteca == 'n' or escolha_biblioteca == 'N':
        print("Saindo...")  # Exibe a mensagem de saída
        sys.exit()  # Encerra o programa

# Verifica as bibliotecas antes de continuar com o resto do programa
verificar_bibliotecas()


#=====================================================================================#

import pandas as pd # Manipula os dados do json
from tkinter import filedialog, messagebox, Tk #
import tkinter as tk
import ast
import json
import os
from datetime import datetime
import matplotlib.pyplot as plt
from collections import Counter, defaultdict
from openpyxl.workbook import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo


def limparTerminal():
    if os.name == 'nt':
        os.system('cls')
    else:  # Para Linux e MacOS
        os.system('clear')


def parser():

    #================ Funç verificar_cabecalho================

    def verificar_cabecalho(arquivo_entrada): #Cria o Parametro arquivo_entrada aonde o arquivo selecionado do usuario será armazenado
        data = pd.read_csv(arquivo_entrada, sep='\t') #data recebe arquivo_entrada lê ele e adiciona um parametro \t que significa T separado por abulação
        #print(data.columns)

    #=========================================================

    #============================ Funç filtra_e_criar_arquivo=========================
    def filtrar_e_criar_arquivo(arquivos_entrada, arquivo_saida):
        data_list = []
        for arquivo_entrada in arquivos_entrada:
            data = pd.read_csv(arquivo_entrada, sep='\t')
            data_list.append(data)
        
        data = pd.concat(data_list, ignore_index=True)
        expected_columns = ['Date Time', 'Customer Name', 'Customer Email', ' Custom Fields', 'Service'] #Id da coluna do .TSV
        
        # Verificar se todas as colunas esperadas estão presentes
        for col in expected_columns:
            if col not in data.columns:
                raise KeyError(f"A coluna {col} não está presente no arquivo de entrada.")
        
        selected_data_filtered = data[expected_columns]
        selected_data_filtered.columns = ['Data', 'Nome', 'email', 'custom_fields', 'Espaco']

        def extract_event_info(x, key):
            try:
                event_info = ast.literal_eval(x)
                if isinstance(event_info, dict):
                    return event_info.get(key, '')
                else:
                    return x
            except (SyntaxError, ValueError):
                return ''

        selected_data_filtered['nome_evento'] = selected_data_filtered['custom_fields'].apply(lambda x: extract_event_info(x, 'Nome do evento'))
        selected_data_filtered['quantidade_participantes'] = selected_data_filtered['custom_fields'].apply(lambda x: extract_event_info(x, 'Quantidade estimada de participantes'))
        
        selected_data_filtered = selected_data_filtered.drop(columns=['custom_fields'])
        
        json_file = arquivo_saida
        selected_data_filtered.to_json(json_file, orient="records", indent=4)
        print(f'O arquivo {json_file} foi criado')
    #=================================================================================

    def selecionar_arquivos():
        root = Tk()
        root.withdraw()  # Fecha a janela principal do Tkinter
        pagArqs = filedialog.askopenfilenames()
        print(f'Estes são os arquivos selecionados: {pagArqs}')
        return pagArqs

    def selecionar_local_para_salvar():
        root = Tk()
        root.withdraw()  # Fecha a janela principal do Tkinter
        pagArq = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        print(f'Esse é o local selecionado para salvar: {pagArq}')
        return pagArq

    arquivos_entrada = selecionar_arquivos()
    limparTerminal()

    verificar_cabecalho(arquivos_entrada[0])  # Verificar o primeiro arquivo antes de seguir

    arquivo_saida = selecionar_local_para_salvar()
    filtrar_e_criar_arquivo(arquivos_entrada, arquivo_saida)
    limparTerminal()

    with open(arquivo_saida, 'r') as file:
        data = json.load(file)

    filtered_data = [item for item in data if item.get('Espaco') not in ['Sala amarela', 'Sala de Reuniao 2']]

    with open('colendario.json', 'w') as file:
        json.dump(filtered_data, file, indent=2)
    print("Processo concluído. Verifique o arquivo: colendario.json")

    with open('colendario.json', 'r') as file:
        data = json.load(file)

    def classificarEvento(hora):
        if 8 <= hora < 12:
            return 'Evento de Manhã'
        elif 12 <= hora < 18:
            return 'Evento de Tarde'
        else:
            return 'Evento Dia Todo'

    eventos_por_dia = {}

    for evento in data:
        try:
            data_evento = datetime.strptime(evento['Data'], '%d/%m/%Y %H:%M').date()
        except ValueError as e:
            print(f"Erro ao analisar a data do evento: {evento['Data']}. Erro: {e}")
            continue

        chave = (data_evento, evento['email'], evento['nome_evento'])

        if chave in eventos_por_dia:
            eventos_por_dia[chave].append(evento)
        else:
            eventos_por_dia[chave] = [evento]

    for eventos in eventos_por_dia.values():
        if len(eventos) > 1:
            for evento in eventos:
                evento['classificacao'] = 'Evento Dia Todo'
        else:
            evento = eventos[0]
            dataHora = datetime.strptime(evento['Data'], '%d/%m/%Y %H:%M')
            evento['classificacao'] = classificarEvento(dataHora.hour)

    with open('arq_filtro_data.json', 'w') as file:
        json.dump(data, file, indent=4)

    limparTerminal() 
    print("Processo concluído. Verifique o arquivo: arq_filtro_data.json")


def analise(pasta_excel, pasta_imagens):
    # Calcula a média de participantes e eventos
    def calcular_media(json_files):
        dados = []
        
        for file in json_files:
            with open(file, 'r') as f:
                json_data = json.load(f)
                for evento in json_data:
                    mes = evento.get("Data", "").split("/")[1] + "/" + evento.get("Data", "").split("/")[2].split(" ")[0]
                    participantes = evento.get("quantidade_participantes", 0)
                    try:
                        participantes = int(participantes) if participantes else 0
                    except ValueError:
                        participantes = 0
                    
                    # Substituir por 30 se "quantidade_participantes" for 0 e "Espaco" for "Palco Principal"
                    if participantes == 0 and evento.get("Espaco") == "Palco Principal":
                        participantes = 30
                    
                    dados.append({
                        "mes": mes,
                        "participantes": participantes,
                        "evento": evento.get("nome_evento", "Desconhecido"),
                        "espaco": evento.get("Espaco", "Desconhecido"),
                        "pessoa": evento.get("Nome", "Desconhecido")
                    })
        
        df = pd.DataFrame(dados)
        
        # Calcular o total e a média de participantes por mês
        dados_mensais = df.groupby("mes").agg(
            total_participantes=pd.NamedAgg(column="participantes", aggfunc="sum"),
            total_eventos=pd.NamedAgg(column="participantes", aggfunc="count"),
            media_participantes=pd.NamedAgg(column="participantes", aggfunc="mean")
        ).reset_index()
        
        # Arredondar a média de participantes para 2 casas decimais
        dados_mensais["media_participantes"] = dados_mensais["media_participantes"].round(2)
        
        # Calcular a média de participantes por evento (juntando eventos com o mesmo nome)
        dados_eventos = df.groupby("evento").agg(
            total_participantes=pd.NamedAgg(column="participantes", aggfunc="sum"),
            total_eventos=pd.NamedAgg(column="participantes", aggfunc="count"),
            media_participantes=pd.NamedAgg(column="participantes", aggfunc="mean")
        ).reset_index()
        
        # Arredondar a média de participantes para 2 casas decimais
        dados_eventos["media_participantes"] = dados_eventos["media_participantes"].round(2)
        
        return df, dados_mensais, dados_eventos

    # Função para plotar os dados
    def plotar_dados(dados_mensais, pasta_imagens):
        plt.figure(figsize=(10, 6))
        plt.plot(dados_mensais["mes"], dados_mensais["total_participantes"], marker='o', label="Total de Participantes")
        plt.plot(dados_mensais["mes"], dados_mensais["total_eventos"], marker='o', label="Total de Eventos")
        plt.plot(dados_mensais["mes"], dados_mensais["media_participantes"], marker='o', label="Média de Participantes")
        
        plt.xlabel("Mês")
        plt.ylabel("Contagem")
        plt.title("Participantes e Eventos Mensais")
        plt.legend()
        plt.grid(True)
        plt.xticks(rotation=45)
        
        plt.savefig(os.path.join(pasta_imagens, "participantes_eventos_mensais.png"))
        plt.show()

    # Função para salvar os dados em um arquivo Excel
    def salvar_para_excel(dados_mensais, dados_eventos, pasta_excel):
        with pd.ExcelWriter(os.path.join(pasta_excel, "participantes_eventos.xlsx")) as writer:
            dados_mensais.to_excel(writer, sheet_name="Mensal", index=False)
            dados_eventos.to_excel(writer, sheet_name="Por Evento", index=False)

    # Função para mostrar gráfico de eventos por mês
    def grafico_eventos_por_mes(df, pasta_imagens):
        eventos_por_mes = df.groupby("mes").size()
        eventos_por_mes.plot(kind='bar', figsize=(10, 6))
        
        plt.xlabel("Mês")
        plt.ylabel("Número de Eventos")
        plt.title("Número de Eventos por Mês")
        
        plt.savefig(os.path.join(pasta_imagens, "eventos_por_mes.png"))
        plt.show()

    # Função para mostrar gráfico de uso de espaços
    def grafico_uso_espacos(df, pasta_imagens):
        uso_espacos = df.groupby("espaco").size()
        uso_espacos.plot(kind='bar', figsize=(10, 6))
        
        plt.xlabel("Espaço")
        plt.ylabel("Número de Eventos")
        plt.title("Uso dos Espaços")
        
        plt.savefig(os.path.join(pasta_imagens, "uso_espacos.png"))
        plt.show()

    # Função para mostrar gráfico de eventos por pessoa
    def grafico_eventos_por_pessoa(df, pasta_imagens):
        eventos_por_pessoa = df.groupby("pessoa").size()
        eventos_por_pessoa.plot(kind='bar', figsize=(10, 6))
        
        plt.xlabel("Pessoa")
        plt.ylabel("Número de Eventos")
        plt.title("Número de Eventos por Pessoa")
        
        plt.savefig(os.path.join(pasta_imagens, "eventos_por_pessoa.png"))
        plt.show()

    # Função para mostrar gráfico de participantes por mês
    def grafico_participantes_por_mes(dados_mensais, pasta_imagens):
        dados_mensais.plot(x='mes', y='total_participantes', kind='line', marker='o', figsize=(10, 6))
        
        plt.xlabel("Mês")
        plt.ylabel("Total de Participantes")
        plt.title("Total de Participantes por Mês")
        
        plt.savefig(os.path.join(pasta_imagens, "participantes_por_mes.png"))
        plt.show()

    # Abrir diálogo de seleção de arquivos JSON
    root = Tk()
    root.withdraw()
    
    json_files = filedialog.askopenfilenames(title="Selecione Arquivos JSON", filetypes=[("Arquivos JSON", "*.json")])
    
    if not json_files:
        print("Nenhum arquivo selecionado.")
        return
    
    # Verificar se as pastas existem, se não, criar as pastas
    if not os.path.exists(pasta_excel):
        os.makedirs(pasta_excel)
        
    if not os.path.exists(pasta_imagens):
        os.makedirs(pasta_imagens)
    
    # Calcular a média de participantes e eventos
    df, dados_mensais, dados_eventos = calcular_media(json_files)
    
    # Salvar os dados em um arquivo Excel
    salvar_para_excel(dados_mensais, dados_eventos, pasta_excel)
    
    while True:
        print("\nMenu:")
        print("1. Mostrar gráfico de eventos por mês")
        print("2. Mostrar gráfico de uso de espaços")
        print("3. Mostrar gráfico de eventos por pessoa")
        print("4. Mostrar gráfico de participantes por mês")
        print("0. Sair")

        escolha = input("\nEscolha uma opção: ")

        if escolha == '1':
            grafico_eventos_por_mes(df, pasta_imagens)
        elif escolha == '2':
            grafico_uso_espacos(df, pasta_imagens)
        elif escolha == '3':
            grafico_eventos_por_pessoa(df, pasta_imagens)
        elif escolha == '4':
            grafico_participantes_por_mes(dados_mensais, pasta_imagens)
        elif escolha == '0':
            break
        else:
            print("Opção inválida. Tente novamente.")

    print("Processo concluído com sucesso.")


def separar():
    # Função para analisar a data e retornar uma tupla de ano e mês
    def analisar_data(data_str):
        data_obj = datetime.strptime(data_str, "%d/%m/%Y %H:%M")
        return data_obj.year, data_obj.month

    # Função para agrupar eventos por ano e mês
    def agrupar_eventos_por_mes_e_ano(eventos):
        eventos_agrupados = defaultdict(lambda: defaultdict(list))
        for evento in eventos:
            ano, mes = analisar_data(evento['Data'])
            eventos_agrupados[ano][mes].append(evento)
        return eventos_agrupados

    # Função para salvar os eventos agrupados em arquivos JSON
    def salvar_eventos_agrupados(eventos_agrupados, pasta_destino):
        for ano, meses in eventos_agrupados.items():
            for mes, eventos in meses.items():
                nome_arquivo = f"eventos_{ano}_{mes:02d}.json"
                caminho_arquivo = os.path.join(pasta_destino, nome_arquivo)
                with open(caminho_arquivo, 'w', encoding='utf-8') as arquivo:
                    json.dump(eventos, arquivo, ensure_ascii=False, indent=2)
                print(f"Arquivo salvo: {caminho_arquivo}")

    # Função para abrir um diálogo de arquivo e processar o arquivo JSON selecionado
    def abrir_arquivo():
        caminho_arquivo = filedialog.askopenfilename(filetypes=[("Arquivos JSON", "*.json")])
        if caminho_arquivo:
            try:
                with open(caminho_arquivo, 'r', encoding='utf-8') as arquivo:
                    dados = json.load(arquivo)
                    eventos_agrupados = agrupar_eventos_por_mes_e_ano(dados)
                    pasta_destino = filedialog.askdirectory(title="Selecione a pasta para salvar os arquivos")
                    if pasta_destino:
                        salvar_eventos_agrupados(eventos_agrupados, pasta_destino)
                        messagebox.showinfo("Sucesso", "Eventos agrupados e salvos com sucesso!")
                    else:
                        messagebox.showwarning("Aviso", "Nenhuma pasta selecionada. Operação cancelada.")
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao processar o arquivo: {e}")

    # Cria a janela principal
    janela = tk.Tk()
    janela.title("Agrupador de Eventos JSON")

    # Cria e posiciona o botão para abrir o diálogo de arquivo
    botao_abrir = tk.Button(janela, text="Abrir Arquivo JSON", command=abrir_arquivo)
    botao_abrir.pack(pady=20)

    # Executa a aplicação
    janela.mainloop()



def converterExcel():
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


limparTerminal()


while True:

    print('===========MENU==========')
    print('1 - Parser: ')
    print('2 - Analise: ')
    print('3 - Separar: ')
    print('4 - Limpar terminal: ')
    print('5 - Converter Excell: ')
    print('0 - Sair: ')
    print('=========================')
    menu = input('Digite sua Escolha: ')

    if menu == '1':    
        print('----- PARSER SELECIONADO ----->')
        parser()
    elif menu == '2':
        print('----- Analise Selecionada ----->')
        pasta_excel = ''#Se quiser mudar aonde o excel e salvo mude nessa linha
        pasta_imagens = ''# Se quiser mudar aonde as imagens e graficos e salvo mude nessa linha

        # Executar a função analise
        analise(pasta_excel, pasta_imagens)
    elif menu == '3':
        print('----- Separa Selecionado ----->')
        separar()
    elif menu == '4':
        print('---- Limpar Terminal ----')
        limparTerminal()
    elif menu == '5':
        print('\n---- Converter Excel ----\n')
        converterExcel()
        input('/\nAperte Enter para continuar...')
        limparTerminal()
    elif menu == '0':  
        limparTerminal()
        print('SAINDO')
        break
    else:
        print('Escolha uma opção valida !!')
        
