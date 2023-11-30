import tkinter as tk
from tkinter import filedialog
import json

def arquivo():
    pagArq = filedialog.askopenfilename()
    print(f'Esse e o arquivo selecionado{pagArq}')
    return pagArq

root=tk.Tk()
root.title("Parser")
btn = tk.Button(root, text='Selecionar arquivo', command=arquivo)
btn.pack()
root.mainloop()

# ABRE ARQUIVO E FAZ A LEITURA  
with open(arquivo(), 'r', encoding='utf-8') as ler:
    data = ler.readlines()

data = [line.strip().split('\t') for line in data]

selected_data = [[row[0], row[2], row[17]] for row in data]

# CRIA O DICIONÁRIO PARA O JSON
json_dict = {"data": []}
for row in selected_data:
    json_dict["data"].append({"col1": row[0], "col2": row[1], "col3": row[2]})

# CRIA O ARQUIVO JSON
json_file = 'data.json'
with open(json_file, 'w', encoding='utf-8') as escrever:
    json.dump(json_dict, escrever, ensure_ascii=False, indent=4)
    print(f'O arquivo {json_file} foi criado')
