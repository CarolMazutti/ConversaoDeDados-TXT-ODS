import pandas as pd
import os
import tkinter as tk
from tkinter import simpledialog, filedialog, messagebox  

# Função para processar o arquivo
def processamento(arquivo_txt):
    # Criar uma janela para solicitar ALIQ e TIPO DE SERVIÇO
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal

    # Solicitar alíquota
    aliq_input = simpledialog.askstring("Entrada de Dados", "Informe a alíquota (sem %):") 
    aliq_input = aliq_input.replace('.', ',')
    aliq = float(aliq_input.replace(',', '.'))

    # Formatar a alíquota para ter duas casas decimais
    aliq_formatado = f"{aliq:.2f}".replace('.', ',')

    # Solicitar tipo de serviço
    tipo_servico = simpledialog.askstring("Entrada de Dados", "Informe o Tipo de Serviço:")

    # Lista para armazenar os dados
    dados = []

    # Ler o arquivo txt
    with open(arquivo_txt, 'r') as file:
        for linha in file:
            # Extrair informações com base nas posições dos caracteres
            cpf_cnpj = linha[0:14].strip()
            data = linha[81:89].strip()
            nf = linha[94:103].strip()
            valor = linha[135:147].strip()
            base = linha[147:159].strip()

            # Converter a data para o formato dd/mm/aaaa
            data_formatada = pd.to_datetime(data, format='%Y%m%d').strftime('%d/%m/%Y')

            # Adicionar os dados extraídos à lista
            dados.append([cpf_cnpj, data_formatada, nf, valor, base])

    # Calcular ICMS e adicionar as novas colunas
    for i in range(len(dados)):
        base_valor = float(dados[i][4]) / 100  # Converter BASE para float
        icms = base_valor * aliq / 100  # Calcular ICMS

        # Adicionar ALIQ, ICMS e TIPO DE SERVIÇO
        dados[i].extend([aliq_formatado, f"{icms:.2f}".replace('.', ','), tipo_servico])

    # Criar um DataFrame com os dados
    df = pd.DataFrame(dados, columns=['CPF/CNPJ', 'DATA', 'NF', 'VALOR', 'BASE', 'ALIQ', 'ICMS', 'TIPO DE SERVIÇO'])

    # Função para salvar o arquivo no formato escolhido
    def salvar_arquivo(formato):
        # Criar uma janela de "Em Processamento"
        processamento_janela = tk.Tk()
        processamento_janela.title("Processando")
        processamento_janela.geometry("200x100")
        label = tk.Label(processamento_janela, text="Em processamento...", padx=20, pady=20)
        label.pack()
        processamento_janela.update()  # Atualiza a janela para exibir a mensagem

        # Salvar o arquivo no formato escolhido
        if formato == 'ods':
            caminho_ods = filedialog.asksaveasfilename(defaultextension=".ods", filetypes=[("ODS files", "*.ods")])
            df.to_excel(caminho_ods, index=False, engine='odf')
            messagebox.showinfo("Sucesso", "Arquivo salvo como ODS com sucesso!")
        elif formato == 'csv':
            caminho_csv = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
            df.to_csv(caminho_csv, index=False, sep=';', encoding='utf-8')
            messagebox.showinfo("Sucesso", "Arquivo salvo como CSV com sucesso!")

        # Fechar a janela de processamento
        processamento_janela.destroy()
        root.quit()  # Finaliza o programa após o salvamento

    # Criar uma nova janela para selecionar o formato de salvamento
    formato_janela = tk.Tk()
    formato_janela.title("Escolha o formato")
    formato_janela.geometry("300x150")

    # Botão para salvar como ODS
    botao_ods = tk.Button(formato_janela, text="Salvar como ODS", command=lambda: [formato_janela.destroy(), salvar_arquivo('ods')])
    botao_ods.pack(pady=10)

    # Botão para salvar como CSV
    botao_csv = tk.Button(formato_janela, text="Salvar como CSV", command=lambda: [formato_janela.destroy(), salvar_arquivo('csv')])
    botao_csv.pack(pady=10)

    formato_janela.mainloop()  # Inicia o loop da janela de seleção de formato

# Criar uma nova janela para solicitar o arquivo .txt
def selecionar_arquivo():

    # Exibir uma mensagem para o usuário
    if messagebox.askokcancel("Escolha o arquivo", "Clique em OK para escolher o arquivo"):
        # Solicitar ao usuário que escolha o arquivo .txt
        arquivo_txt = filedialog.askopenfilename(title="Selecione o arquivo", filetypes=[("Text Files", "*.001")])

        # Verificar se o arquivo .txt existe antes de processar
        if not os.path.exists(arquivo_txt):
            messagebox.showerror("Erro", f"O arquivo {arquivo_txt} não foi encontrado.")
        else:
            # Executar a função de processamento
            processamento(arquivo_txt)

selecionar_arquivo()  # Chamar a função para iniciar o processo de seleção de arquivo