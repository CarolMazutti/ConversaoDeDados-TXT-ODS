import pandas as pd
import os
import tkinter as tk
from tkinter import simpledialog

# Caminhos para os arquivos de entrada e saída
arquivo_txt = '\\\\192.168.1.142\\Geral\\PROGRAMAS\\Importações\\dados_cliente.001'
arquivo_ods = '\\\\192.168.1.142\\Geral\\PROGRAMAS\\Importações\\dados_tratados.ods'

# Verificar se o arquivo .txt existe antes de processar
if not os.path.exists(arquivo_txt):
    print(f"Erro: O arquivo {arquivo_txt} não foi encontrado.")
else:
    # Função para ler o arquivo .txt e extrair dados com base nas posições de caracteres
    def processamento(arquivo_txt, arquivo_ods):
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

        processamento_janela = tk.Tk()
        processamento_janela.title("Processando")
        processamento_janela.geometry("200x100")
        label = tk.Label(processamento_janela, text="Em processamento...", padx=20, pady=20)
        label.pack()
        processamento_janela.update()  # Atualiza a janela para exibir a mensagem
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

            # Formatar VALOR e BASE para incluir a vírgula antes dos dois últimos dígitos
            valor_formatado = f"{float(dados[i][3]) / 100:.2f}".replace('.', ',')
            base_formatado = f"{base_valor:.2f}".replace('.', ',')
            icms_formatado = f"{icms:.2f}".replace('.', ',')

            # Adicionar ALIQ, ICMS e TIPO DE SERVIÇO
            dados[i].extend([aliq_formatado, icms_formatado, tipo_servico])

        # Criar um DataFrame com os dados
        df = pd.DataFrame(dados, columns=['CPF/CNPJ', 'DATA', 'NF', 'VALOR', 'BASE', 'ALIQ', 'ICMS', 'TIPO DE SERVIÇO'])

        # Formatar VALOR e BASE para incluir a vírgula antes dos dois últimos dígitos
        df['VALOR'] = df['VALOR'].astype(float).map(lambda x: f"{x / 100:.2f}".replace('.', ','))
        df['BASE'] = df['BASE'].astype(float).map(lambda x: f"{x / 100:.2f}".replace('.', ','))
        df['ALIQ'] = df['ALIQ'].astype(str).str.replace('.', ',')
        df['ICMS'] = df['ICMS'].astype(str).str.replace('.', ',')

        # Salvar o DataFrame em um arquivo .ods
        df.to_excel(arquivo_ods, engine='odf', index=False)

        processamento_janela.destroy()

    # Executar a função
    processamento(arquivo_txt, arquivo_ods)