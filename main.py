import tkinter as tk
from tkinter import filedialog
from docx import Document
import pandas as pd
import os

def preencher_certificado(template_path, output_path, data):
    doc = Document(template_path)

    for paragraph in doc.paragraphs:
        for key, value in data.items():
            if key in paragraph.text:
                for run in paragraph.runs:
                    run.text = run.text.replace(key, str(value))  # Convertendo value para string
    doc.save(output_path)

def gerar_certificados(xlsx_path, template_path):
    # Convertendo o arquivo xlsx para csv temporário
    csv_temp_path = 'temp.csv'
    df = pd.read_excel(xlsx_path)
    df.to_csv(csv_temp_path, sep=';', index=False)

    # Lendo o arquivo CSV temporário e gerando os certificados
    df_csv = pd.read_csv(csv_temp_path, sep=';')
    for idx, row in df_csv.iterrows():
        data ={
            '[NOME]': row['NOME'],
            '[SOBRENOME]': row['SOBRENOME'],
            '[EVENTO]': row['EVENTO'],
            '[DATA]': row['DATA'],
            '[CARGA_HORARIA]': row['CARGA_HORARIA']
        }

        output_path = f'Certificado_{row["EVENTO"]}_{row["NOME"]}_{row["SOBRENOME"]}.docx'
        preencher_certificado(template_path, output_path, data)

    # Removendo o arquivo csv temporário
    os.remove(csv_temp_path)
    tk.messagebox.showinfo("Concluído", "Certificados gerados com sucesso!")

def selecionar_arquivo_xlsx():
    xlsx_path = filedialog.askopenfilename(title="Selecione o arquivo xlsx")
    entry_xlsx.delete(0, tk.END)
    entry_xlsx.insert(0, xlsx_path)

def selecionar_arquivo_template():
    template_path = filedialog.askopenfilename(title="Selecione o modelo de certificado")
    entry_template.delete(0, tk.END)
    entry_template.insert(0, template_path)

def gerar_certificados_interface():
    xlsx_path = entry_xlsx.get()
    template_path = entry_template.get()
    gerar_certificados(xlsx_path, template_path)

# Criando a janela principal
root = tk.Tk()
root.title("Gerador de Certificados")

# Criando os widgets
label_xlsx = tk.Label(root, text="Selecione o arquivo xlsx:")
label_template = tk.Label(root, text="Selecione o modelo de certificado:")
entry_xlsx = tk.Entry(root, width=50)
entry_template = tk.Entry(root, width=50)
button_browse_xlsx = tk.Button(root, text="Procurar", command=selecionar_arquivo_xlsx)
button_browse_template = tk.Button(root, text="Procurar", command=selecionar_arquivo_template)
button_generate = tk.Button(root, text="Gerar Certificados", command=gerar_certificados_interface)

# Posicionando os widgets na janela
label_xlsx.grid(row=0, column=0, padx=10, pady=5, sticky="e")
entry_xlsx.grid(row=0, column=1, padx=10, pady=5)
button_browse_xlsx.grid(row=0, column=2, padx=10, pady=5)

label_template.grid(row=1, column=0, padx=10, pady=5, sticky="e")
entry_template.grid(row=1, column=1, padx=10, pady=5)
button_browse_template.grid(row=1, column=2, padx=10, pady=5)

button_generate.grid(row=2, column=1, pady=10)

# Iniciando o loop da interface gráfica
root.mainloop()

