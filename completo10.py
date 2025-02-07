import os
import win32com.client
from datetime import datetime, timedelta, timezone
import pytz
import re
import pdfplumber
import shutil
import pandas as pd
from pathlib import Path
import time 
import sys

caminho_base = os.getcwd()  # Obtém o diretório atual

caminho = os.path.join(caminho_base, 'PASTA LCD A PAGAR')
destino = os.path.join(caminho_base, 'RECIBO DE PAGAMENTO')  # Junta o caminho com a pasta desejada
local_pendencia_recibo = os.path.join(caminho_base, 'PENDENCIA DE RECIBO')  # Junta o caminho com a pasta desejada
caminho_temporia = os.path.join(caminho_base, 'TEMPORARIA')

# Caminho base para a pasta DESCARGAS
caminho_descargas = r'G:\Meu Drive\DESCARGAS'
caminho_comprovantes = os.path.join(caminho_descargas, "COMPROVANTES DE PAGTO LCD")
caminho_enviados = os.path.join(caminho_descargas, "COMPROVANTES ENVIADOS")
caminho_descarga_lcd = os.path.join(caminho_descargas, 'PASTA LCD A PAGAR')
caminho_descarga_recibo = os.path.join(caminho_descargas, 'RECIBO DE PAGAMENTO')  # Junta o caminho com a pasta desejada
caminho_descarga_pendencia_recibo = os.path.join(caminho_descargas, 'PENDENCIA DE RECIBO')  
caminho_pasta_error = os.path.join(caminho_descargas, 'ERROR')

local_recibo = ''

#1 PARTE

try:
    caminho_planilha = os.path.join(caminho_descarga_lcd, "PENDENCIAS_DESCARGAS.xlsx")
    df = pd.read_excel(caminho_planilha, dtype=str)
except:
    pass

def processar_pdf_comprovante(pdf_path):
    # Função para extrair informações do PDF
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            texto = page.extract_text()
            linhas = texto.split("\n")
            dados = {"Página": page_num}
            
            for i, linha in enumerate(linhas):

                if "Doc Contas a Pagar" in linha:
                    # Remover "Doc Contas a Pagar" e trabalhar com o restante da linha
                    doc_contas = linha.replace("Doc Contas a Pagar", "").strip()

                    # Padrão para capturar o número desejado
                    padrao = r'\b\d+-\d{1,3}\.\d{3}-\d\b'

                    # Usando re.search para encontrar o padrão
                    match = re.search(padrao, doc_contas)

                    if match:
                        resultado = match.group()
                        resultado = resultado.replace('.','')
                        resultado = resultado
                        lcd = resultado
                    else:
                        print("Nenhum número encontrado.")
                        dados["CPG"] = None
                    
    return lcd
# Listar todos os arquivos na pasta
arquivos = os.listdir(caminho_comprovantes)

# Filtrar apenas arquivos PDF
pdfs = [arquivo for arquivo in arquivos if arquivo.endswith('.pdf')]
qtd_pdf = len(pdfs)
i = 0

# Conectar-se ao Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
caixa_entrada = namespace.GetDefaultFolder(6)  # 6 é a caixa de entrada

# Definir a data limite (uma semana atrás) com fuso horário UTC
data_limite = datetime.now(pytz.UTC) - timedelta(days=1)

# Filtrar e-mails até uma semana atrás
itens = caixa_entrada.Items
itens.Sort("[ReceivedTime]", True)  # Ordenar os e-mails pela data de recebimento

# Iterar pelos e-mails
for item in itens:
    # Garantir que o ReceivedTime tenha o mesmo fuso horário
    if item.ReceivedTime.tzinfo is None:  # Se o e-mail não tem fuso horário
        item.ReceivedTime = item.ReceivedTime.replace(tzinfo=pytz.UTC)

    # Comparar a data
    if item.ReceivedTime < data_limite:
        break  # Parar quando o e-mail for mais antigo que a data limite
    
    # Verificar se o e-mail possui anexos
    if item.Attachments.Count > 0:
        comprovantes_enviados = []
        for anexo in item.Attachments:
            # Salvar o anexo temporariamente
            caminho_anexo = os.path.join(caminho_temporia, anexo.FileName)
            anexo.SaveAsFile(caminho_anexo)
            
            # Verificar se o anexo é um PDF
            if caminho_anexo.endswith('.pdf'):
                # Extração do número do comprovante (exemplo: o número é o nome do arquivo PDF)
                numero_comprovante_email = anexo.FileName.replace('.pdf', '')
                
                # Verificar se o nome do anexo corresponde a um PDF na pasta
                for pdf in pdfs:
                    
                    numero_comprovante_pasta = pdf.replace('.pdf', '')
                    # Expressão regular para capturar o texto entre "LCD" e "-PC"
                    padrao = r"LCD (.*?)\-PC"

                    # Buscar a correspondência
                    resultado = re.search(padrao, numero_comprovante_pasta)

                    # Verificar se encontrou o padrão
                    if resultado:
                        # Extrair o texto encontrado
                        extracao = resultado.group(1)

                    else:
                        print("Não foi possível encontrar o padrão.")
                    local = os.path.join(caminho_temporia, anexo.FileName)
                    
                    try:
                        lcd = processar_pdf_comprovante(local)
                    except:
                        lcd = None
                    
                    if lcd == extracao:
                        if i == qtd_pdf:
                            break

                        local = os.path.join(caminho_comprovantes, pdf)

                        #Verifica se o número LCD está na coluna "Nº LCD" da planilha
                        if extracao in df["Nº LCD"].values:

                            # Remove a linha correspondente
                            df = df[df["Nº LCD"] != extracao]

                            # Salva a planilha atualizada
                            df.to_excel(caminho_planilha, index=False)

                            print(f"✅ Removido {extracao} da planilha.")

                        i = i + 1

                        caminho_pdf = os.path.join(caminho_comprovantes, pdf)
                
                        #Criar uma pasta com a data de hoje (formato YYYY-MM-DD)
                        data_hoje = datetime.now().strftime('%Y-%m-%d')
                        caminho_pasta_data = os.path.join(caminho_enviados, data_hoje)

                        # Criar a pasta da data de hoje se não existir
                        if not os.path.exists(caminho_pasta_data):
                            os.makedirs(caminho_pasta_data)
                            print(f"Pasta criada: {caminho_pasta_data}")
                        
                        # Criar uma subpasta com o nome da LCD dentro da pasta do dia
                        caminho_pasta_lcd = os.path.join(caminho_pasta_data, lcd)

                        # Criar a subpasta LCD se não existir
                        if not os.path.exists(caminho_pasta_lcd):
                            os.makedirs(caminho_pasta_lcd)
                            print(f"Pasta da LCD criada: {caminho_pasta_lcd}")
                        
                        # #Caminho de destino para o arquivo PDF dentro da subpasta LCD
                        destino_pdf = os.path.join(caminho_pasta_lcd, pdf)
                        comprovantes_enviados.append(destino_pdf)

                        # #Mover o arquivo para a nova pasta
                        shutil.move(caminho_pdf, destino_pdf)

                        print(f"Arquivo {caminho_pdf} movido para {caminho_pasta_lcd}")

                        # Listar todos os arquivos na pasta
                        arquivos = os.listdir(caminho_descarga_lcd)

                        # Filtrar apenas arquivos PDF
                        pdfs_lcd = [arquivo for arquivo in arquivos if arquivo.endswith('.pdf')]

                        for pdf_lcd in pdfs_lcd:
                            numero_lcd_pasta = pdf_lcd.replace('.pdf', '')
                            # Expressão regular para capturar o texto entre "LCD" e "-PC"
                            padrao = r"LCD (.*?)\-PC"

                            # Buscar a correspondência
                            resultado = re.search(padrao, numero_lcd_pasta)

                            # Verificar se encontrou o padrão
                            if resultado:
                                # Extrair o texto encontrado
                                extracao_lcd = resultado.group(1)

                                if lcd == extracao_lcd:
                                    print('achou lcd')
                                    caminho_pdf_lcd = os.path.join(caminho_descarga_lcd, pdf_lcd)
                                    destino_pdf_lcd = os.path.join(caminho_pasta_lcd, pdf_lcd)
                                    shutil.move(caminho_pdf_lcd, destino_pdf_lcd)
                                    print(f"Arquivo {caminho_pdf_lcd} movido para {caminho_pasta_lcd}")

                        # Listar todos os arquivos na pasta
                        arquivos = os.listdir(caminho_descarga_recibo)

                        # Filtrar apenas arquivos PDF
                        pdfs_recibo = [arquivo for arquivo in arquivos if arquivo.endswith('.pdf')]

                        for pdf_recibo in pdfs_recibo:
                            numero_recibo_pasta = pdf_recibo.replace('.pdf', '')
                            # Expressão regular para capturar o texto entre "LCD" e "-PC"
                            padrao = r"LCD (.*?)\-PC"

                            # Buscar a correspondência
                            resultado = re.search(padrao, numero_recibo_pasta)

                            # Verificar se encontrou o padrão
                            if resultado:
                                # Extrair o texto encontrado
                                extracao_recibo = resultado.group(1)

                                if lcd == extracao_recibo:
                                    print('achou recibo')
                                    caminho_pdf_recibo = os.path.join(caminho_descarga_recibo, pdf_recibo)
                                    destino_pdf_recibo = os.path.join(caminho_pasta_lcd, pdf_recibo)
                                    shutil.move(caminho_pdf_recibo, destino_pdf_recibo)
                                    print(f"Arquivo {caminho_pdf_recibo} movido para {caminho_pasta_lcd}")

        # Responder ao e-mail com todos os comprovantes encontrados
        if comprovantes_enviados:
            resposta = item.Reply()
            resposta.Subject = "Comprovantes de Pagamento"
            resposta.Body = "Segue em anexo os comprovantes de pagamento."

            # Adicionar todos os comprovantes como anexos
            for comprovante in comprovantes_enviados:
                resposta.Attachments.Add(comprovante)

            resposta.Send()

            print("✅ E-mail enviado com os comprovantes.")

# Percorre todos os arquivos na pasta de origem
for arquivo in os.listdir(caminho_comprovantes):
    caminho_arquivo = os.path.join(caminho_comprovantes, arquivo)  # Caminho completo do arquivo

    # Verifica se é um arquivo (para evitar mover pastas)
    if os.path.isfile(caminho_arquivo):
        destino_arquivo = os.path.join(caminho_pasta_error, arquivo)  # Caminho completo no destino
        shutil.move(caminho_arquivo, destino_arquivo)  # Move o arquivo
        print(f"Movido: {arquivo} -> {caminho_pasta_error}")

print("✅ Todos os arquivos foram movidos com sucesso!")

# Listar todos os arquivos na pasta
arquivos = os.listdir(caminho_temporia)

# Filtrar e excluir os arquivos PDF
for arquivo in arquivos:
    if arquivo.endswith('.pdf'):
        caminho_arquivo = os.path.join(caminho_temporia, arquivo)
        os.remove(caminho_arquivo)  # Excluir o arquivo PDF
        print(f"Arquivo {arquivo} excluído.")


#------------------------------------------------------------------------------------------------------------------------------------------
#2 PARTE

def processar_pdf(pdf_path,tipo_documento,local_recibo,remetente,hora_recebimento,email,agora, tempo_limite):
    # Função para extrair informações do PDF
    with pdfplumber.open(pdf_path) as pdf:
        resultados = []
        for page_num, page in enumerate(pdf.pages, start=1):
            texto = page.extract_text()
            linhas = texto.split("\n")
            dados = {"Página": page_num}
            
            for i, linha in enumerate(linhas):

                if "Registro:" in linha:
                    doc_contas = linha.split(":")
                    if len(doc_contas) > 2:
                        registro = doc_contas[2].strip()
                        uni = registro[:1]
                        registro = registro[4:].replace('.','')
                        dados["Registro"] = registro
                        
                if "Unidade:" in linha:
                    doc_contas = linha.split(":")
                    if uni == '1' or uni == '2' or uni == '3':
                        unidade = doc_contas[1][:19]
                    elif uni == '5':
                        unidade = doc_contas[1][:16]
                    unidade = unidade.lstrip()
                    if len(doc_contas) > 2:
                        dados["Unidade"] = unidade.strip()

                if "Doc Contas a Pagar" in linha:
                    # Remover "Doc Contas a Pagar" e trabalhar com o restante da linha
                    doc_contas = linha.replace("Doc Contas a Pagar", "").strip()

                    # Padrão para capturar o número desejado
                    padrao = r'\b\d+-\d{1,3}\.\d{3}-\d\b'

                    # Usando re.search para encontrar o padrão
                    match = re.search(padrao, doc_contas)

                    if match:
                        resultado = match.group()
                        resultado = resultado.replace('.','')
                        dados["CPG"] = resultado
                    else:
                        print("Nenhum número encontrado.")
                        dados["CPG"] = None
                    
                    # Capturar a data no formato DD/MM/YY
                    padrao_data = r'\b\d{2}/\d{2}/\d{2}\b'
                    match_data = re.search(padrao_data, doc_contas)

                    if match_data:
                        dados["Data"] = match_data.group()

                        # Agora que temos a data, pegamos o número antes dela
                        texto_antes_data = doc_contas[:match_data.start()].strip()
                        
                        # Regex para capturar o número antes da data (40 ou 140)
                        padrao_numero = r'(\d{2,3})$'  # Número no final do trecho antes da data
                        match_numero = re.search(padrao_numero, texto_antes_data)

                        if match_numero:
                            dados["Numero"] = match_numero.group()
                    else:
                        dados["Data"] = None
                        dados["Numero"] = None

                    # Extrair se é REEMBOLSÁVEL ou NÃO REEMBOLSÁVEL
                    if "NÃO REEMBOLSÁVEL" in doc_contas:
                        dados["Reembolsável"] = "NÃO REEMBOLSÁVEL"
                    elif "REEMBOLSÁVEL" in doc_contas:
                        dados["Reembolsável"] = "REEMBOLSÁVEL"
                                        
                # Procurar a linha contendo "Valor:"
                if "Valor:" in linha:
                    valor = re.search(r"\d+,\d{2}", linha)
                    if valor:
                        dados["Valor"] = valor.group()
                
                if "Favorecido:" in linha:
                    favorecido_partes = linha.split(":")
                    if len(favorecido_partes) > 1:
                        # Parte após o "Favorecido:"
                        texto_completo = favorecido_partes[1].strip()

                        # Extrair apenas o primeiro conjunto de números antes de um espaço
                        match = re.match(r'([\d.]+)', texto_completo)
                        if match:
                            apenas_numeros = match.group(1)
                        else:
                            apenas_numeros = ""

                        # Remover os números iniciais para obter o texto restante
                        apenas_texto = texto_completo[len(apenas_numeros):].strip()

                        # Salvar os dados
                        dados["COD PARCEIRO COMERCIAL"] = apenas_numeros
                        dados["NOME PARCEIRO COMERCIAL"] = apenas_texto
                
                # Procurar a linha contendo "Banco:"
                if "Banco:" in linha:
                    banco = re.search(r"Banco:\s*(\d+)", linha)
                    if banco:
                        dados["Banco"] = banco.group(1)
                
                # Procurar a linha contendo "Agência:"
                if "Agência:" in linha:
                    agencia = re.search(r"Agência:\s*(\d+)", linha)
                    if agencia:
                        dados["Agência"] = agencia.group(1)
                
                # Procurar a linha contendo "Conta:"
                if "Conta:" in linha:
                    conta = re.search(r"Conta:\s*([\w\-]+)", linha)
                    chave = re.search(r"Chave:\s*([A-Za-z0-9\-\/]+)", linha)
                    if conta:
                        dados["Conta"] = conta.group(1)
                    if chave:
                        dados["Tipo de Chave"] = chave.group(1)
                
                # Procurar a linha contendo "CC"
                if "CC:" in linha:
                    doc_contas = linha.split(":")
                    if len(doc_contas) > 1:
                        dados["CC"] = doc_contas[1].strip()

                # Procurar a linha contendo "Chave"
                if "Chave:" in linha:
                    doc_contas = linha.split(":")
                    if len(doc_contas) > 1:
                        dados["Chave"] = doc_contas[1].strip()

            # Adicionar os dados extraídos aos resultados
            resultados.append(dados)   

    #Exibir os resultados extraídos
    for resultado in resultados:
        print(f'tipo de documento:{tipo_documento}')
        
        cod_parceiro = str(resultado['COD PARCEIRO COMERCIAL'])
        cod_parceiro = cod_parceiro.replace(',', '').replace('.', '')

        #LCD NAO EFETUADA VAI PEDIR PARA EFETUAR E NAO SERA ENVIADA PARA O FINANCEIRO
        if tipo_documento == 'LCD' and resultado.get('CPG') is None and resultado.get('Registro') is not None:

            print('nao efetuado')
            #Criar a resposta
            resposta = email.Reply()  
            resposta.Subject = "Re: " + email.Subject  # Mantém o assunto original
            resposta.Body = (
                f"Olá {email.SenderName},\n\n"
                "Pagamento não realizado devido a LCD nao estar efetuada.\n"
                "Favor efetuar e enviar um novo e-mail.\n\n"
                "Atenciosamente,\nFinanceiro"
            )

            # Adicionar alguém em cópia (CC)
            resposta.CC = 'pendencias.financeiro@jettatransportes.com.br'

            # Enviar a resposta
            resposta.Send()

            print("Resposta enviada com sucesso!")

            #Verifica se o arquivo existe antes de excluir
            if os.path.exists(pdf_path):
                os.remove(pdf_path)
                print("Arquivo excluído com sucesso!")
            else:
                print("Arquivo não encontrado.")

            continue
        else:
            print('nada a excluir')

        #CASO ESTA COM A TAREFA ERRADA VAI PEDIR PARA CORRIGIR E NAO SERA ENVIADA PARA O FINANCEIRO
        if resultado['Numero'] != '40':
            print(f'tarefa:{resultado['Numero']} corrigir')
            #Criar a resposta
            resposta = email.Reply()  
            resposta.Subject = "Re: " + email.Subject  # Mantém o assunto original
            resposta.Body = (
                f"Olá {email.SenderName},\n\n"
                "Pagamento não realizado devido a LCD estar com a tarefa errada.\n"
                "Favor corrigir para a tarefa 40 e reenviar.\n\n"
                "Atenciosamente,\nFinanceiro"
            )

            # Adicionar alguém em cópia (CC)
            resposta.CC = 'pendencias.financeiro@jettatransportes.com.br'

            # Enviar a resposta
            resposta.Send()

            print("Resposta enviada com sucesso!")

            #Verifica se o arquivo existe antes de excluir
            if os.path.exists(pdf_path):
                os.remove(pdf_path)
                print("Arquivo excluído com sucesso!")
            else:
                print("Arquivo não encontrado.")

            print("-" * 50)
            continue
            
        else:
            print(f'tarefa:{resultado['Numero']}')

        if tipo_documento == 'LCD':
            recibo = 'NAO'
            nome_arquivo = 'LCD ' + resultado['CPG'] + '-PC ' + cod_parceiro + '-' + resultado['Reembolsável'] + '-CC' + resultado['CC']

            # Caminho do arquivo atual e o novo nome
            
            caminho_novo = os.path.join(caminho, f"{nome_arquivo}.pdf")
            
            try:
                os.rename(pdf_path, caminho_novo)
                print(f"Arquivo renomeado para: {caminho_novo}")
            except FileNotFoundError:
                print("Arquivo não encontrado. Verifique o caminho.")
            except PermissionError:
                print("Permissão negada. Verifique se o arquivo não está aberto.")
            except Exception as e:
                print(f"Ocorreu um erro: {e}")

            print("-" * 50)

        # Supondo que resultado['Chave'] tenha o valor 36931887000189
        lcd = resultado['CPG']
        pix = int(resultado['Chave'])
        
        # Suponha que você tenha essas variáveis
        dados = {
            "Nº LCD": str(lcd),
            "VALOR LCD": resultado['Valor'],
            "Nº PARCEIRO COML.": cod_parceiro,
            "NOME PARCEIRO COMERCIAL": resultado["NOME PARCEIRO COMERCIAL"],
            "TIPO PIX": resultado['Tipo de Chave'],
            "PIX": pix,  # PIX como número
            "CC": resultado['CC'],
            "SOLICITANTE": remetente,
            "HORARIO": hora_recebimento,
            "REEMBOLSAVEL?": resultado['Reembolsável'],
            "POSSUI RECIBO?": recibo
        }

        nome = f'GRADE {tempo_limite.strftime("%H-%M")} a {agora.strftime("%H-%M")}.xlsx'

        # Caminho do arquivo Excel
        caminho_nova_planilha = os.path.join(caminho_descargas, nome)

        # Verificar se o arquivo já existe
        if os.path.exists(caminho_nova_planilha):
            # Se o arquivo existir, carregar a planilha existente
            df1 = pd.read_excel(caminho_nova_planilha)
            # Adicionar a nova linha de dados ao DataFrame existente
            df1 = pd.concat([df1, pd.DataFrame([dados])], ignore_index=True)
        else:
            # Se o arquivo não existir, criar um novo DataFrame
            df1 = pd.DataFrame([dados])

        # Garantir que a coluna "PIX" seja numérica
        df1['PIX'] = pd.to_numeric(df1['PIX'], errors='coerce')

        # Remover duplicados baseado na coluna 'Nº LCD'
        df1 = df1.drop_duplicates(subset='Nº LCD', keep='last')

        # Salvar o DataFrame atualizado no arquivo Excel com o openpyxl
        with pd.ExcelWriter(caminho_nova_planilha, engine='openpyxl') as writer:
            df1.to_excel(writer, index=False)

            # Usando openpyxl para formatar a coluna "PIX" para evitar notação científica
            workbook = writer.book
            worksheet = workbook.active
            for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=6, max_col=6):  # Coluna PIX (7ª coluna)
                for cell in row:
                    cell.number_format = '0'  # Formato numérico inteiro

        print(f"Planilha atualizada com sucesso: {caminho_nova_planilha}")

        # Caminho do arquivo Excel
        arquivo_excel = os.path.join(caminho_descarga_lcd, 'PENDENCIAS_DESCARGAS.xlsx')

        # Verificar se o arquivo já existe
        if os.path.exists(arquivo_excel):
            # Se o arquivo existir, carregar a planilha existente
            df = pd.read_excel(arquivo_excel)
            # Adicionar a nova linha de dados ao DataFrame existente
            df = pd.concat([df, pd.DataFrame([dados])], ignore_index=True)
        else:
            # Se o arquivo não existir, criar um novo DataFrame
            df = pd.DataFrame([dados])

        # Garantir que a coluna "PIX" seja numérica
        df['PIX'] = pd.to_numeric(df['PIX'], errors='coerce')

        # Remover duplicados baseado na coluna 'Nº LCD'
        df = df.drop_duplicates(subset='Nº LCD', keep='last')

        # Salvar o DataFrame atualizado no arquivo Excel com o openpyxl
        with pd.ExcelWriter(arquivo_excel, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

            # Usando openpyxl para formatar a coluna "PIX" para evitar notação científica
            workbook = writer.book
            worksheet = workbook.active
            for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=6, max_col=6):  # Coluna PIX (7ª coluna)
                for cell in row:
                    cell.number_format = '0'  # Formato numérico inteiro

        print(f"Planilha atualizada com sucesso: {arquivo_excel}")
    return 

def processar_emails_ultimos_10_minutos(pasta_anexos):
    # Conexão com o Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    caixa_de_entrada = outlook.GetDefaultFolder(6)  # 6 = Caixa de entrada

    # Obter o tempo limite (últimos 10 minutos)
    agora = datetime.now()
    tempo_limite = agora - timedelta(minutes=10)

    print(f"Procurando e-mails recebidos nos últimos 10 minutos ({tempo_limite.strftime('%H:%M:%S')} até {agora.strftime('%H:%M:%S')})...\n")
    
    for email in caixa_de_entrada.Items:
        try:
            # Remover o fuso horário de ReceivedTime, convertendo para naive
            data_envio = email.ReceivedTime.replace(tzinfo=None)

            if data_envio >= tempo_limite:
                remetente = email.SenderName
                # hora_recebimento = email.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
                hora_recebimento = email.ReceivedTime.strftime("%d/%m/%Y %H:%M:%S")
                print(f"E-mail encontrado: {email.Subject}")
                print(f"Remetente: {remetente}")
                print(f"Hora de recebimento: {hora_recebimento}\n")
                
                # Extrair informações do corpo do e-mail
                corpo = email.Body
                
                # print(f"Corpo do e-mail:\n{corpo}\n{'-' * 50}")
                if email.Attachments.Count == 0:
                    print("Este e-mail não contém anexos. Enviando resposta...")

                    # Criar a resposta
                    resposta = email.Reply()  
                    resposta.Subject = "Re: " + email.Subject  # Mantém o assunto original
                    resposta.Body = (
                        f"Olá {email.SenderName},\n\n"
                        "Pagamento não realizado devido a falta de procedimento padrão definido(DOCUMENTO NÃO ANEXADO).\n"
                        "Favor corrigir  e anexar o arquivo.\n\n"
                        "Atenciosamente,\nFinanceiro"
                    )

                    # Adicionar alguém em cópia (CC)
                    resposta.CC = 'pendencias.financeiro@jettatransportes.com.br'

                    # Enviar a resposta
                    resposta.Send()

                    print("Resposta enviada com sucesso!")
                    
                else:

                    for anexo in email.Attachments:
                        try:
                            print(f"Anexo encontrado: {anexo.FileName}")
                            caminho_anexo = os.path.join(pasta_anexos, anexo.FileName)
                            anexo.SaveAsFile(caminho_anexo)
                            print(f"Anexo salvo em: {caminho_anexo}")
                            
                            # Verificar se o anexo é um PDF
                            if anexo.FileName.lower().endswith(".pdf"):
                                tipo_documento = 'LCD'
                                processar_pdf(caminho_anexo,tipo_documento,local_recibo,remetente,hora_recebimento,email,agora, tempo_limite)
                        except Exception as e:
                            print(f"Erro ao processar o anexo {anexo.FileName}: {e}")

        except Exception as e:
            print(f"Erro ao processar e-mail: {e}")
    
    return agora, tempo_limite

# Certifique-se de que o diretório existe
os.makedirs(caminho, exist_ok=True)

# Executar o processo
agora, tempo_limite = processar_emails_ultimos_10_minutos(caminho)

#AJUSTAR O NOME DA LCD
print("-" * 50)

# Lista todos os arquivos na pasta e filtra os PDFs que não começam com 'LCD'
pdfs_nao_lcd = [f for f in os.listdir(caminho) if f.endswith(".pdf") and not f.startswith("LCD")]
pdfs_lcd = [f for f in os.listdir(caminho) if f.endswith(".pdf") and f.startswith("LCD")]

# Exibe os arquivos encontrados

for recibo in pdfs_nao_lcd:
    
    # Verificar se o anexo é um PDF
    num_recibo = recibo.replace('.pdf', '').replace('recibo ', '').replace('RECIBO ', '').replace('Recibo ', '') 
    local_recibo = os.path.join(caminho, recibo)

    # Verificar se o nome do anexo corresponde a um PDF na pasta
    if len(pdfs_lcd) == 0:
        print('procurar pasta enviados')
        print('começou')

        # Nome formatado que queremos encontrar
        nome_procurado = recibo.replace("recibo ", "").replace(".pdf", "")
        print(f"Nome procurado: {nome_procurado}")

        for pasta_externa in os.listdir(caminho_enviados):
            caminho_pasta_externa = os.path.join(caminho_enviados, pasta_externa)

            # Verifica se é uma pasta
            if os.path.isdir(caminho_pasta_externa):
                # Agora percorremos as subpastas dentro dessa pasta
                for pasta_interna in os.listdir(caminho_pasta_externa):
                    caminho_pasta_interna = os.path.join(caminho_pasta_externa, pasta_interna)

                    # Se for uma pasta e o nome procurado estiver nela, movemos o arquivo
                    if os.path.isdir(caminho_pasta_interna) and nome_procurado in pasta_interna:

                        # Lista os arquivos que começam com "COMPROV" e terminam com ".pdf"
                        comprovantes = [f for f in os.listdir(caminho_pasta_interna) if f.lower().startswith('comprov') and f.lower().endswith('.pdf')]
                        comprovantes = comprovantes[0].replace("COMPROV ", "RECIBO ")
                        # Exibe os arquivos encontrados
                        comprovantes = os.path.join(caminho_base, comprovantes)
                        print(comprovantes)
                        print(local_recibo)
                        os.rename(local_recibo, comprovantes)
                        destino = os.path.join(caminho_pasta_interna, os.path.basename(comprovantes))
                        shutil.move(comprovantes, destino)
                        print(f"Arquivo movido para: {destino}")

                        # Percorre os arquivos na pasta e verifica quais começam com "LCD" e terminam com ".pdf"
                        arquivos_lcd = [arquivo for arquivo in os.listdir(caminho_pasta_interna) if arquivo.startswith("LCD") and arquivo.endswith(".pdf")]

                        # Exibir os arquivos encontrados
                        if arquivos_lcd:
                            print("Arquivos encontrados:", arquivos_lcd)
                            nome_arquivo = arquivos_lcd[0]  # Salva o primeiro encontrado
                            print("Nome do primeiro arquivo encontrado:", nome_arquivo)
                            lcd_apagar = os.path.join(caminho_descarga_pendencia_recibo, nome_arquivo)

                            #Verifica se o arquivo realmente existe antes de tentar excluí-lo
                            if os.path.exists(lcd_apagar):
                                os.remove(lcd_apagar)  # Exclui o arquivo
                                print(f"✅ Arquivo '{nome_arquivo}' excluído com sucesso!")
                            else:
                                print(f"⚠️ Arquivo '{nome_arquivo}' não encontrado para exclusão.")
                        else:
                            print("Nenhum arquivo encontrado com 'LCD' no início.")
                            # Carregar a planilha onde deseja remover a linha

                        padrao = r"LCD (.*?)\-PC"

                        # Buscar a correspondência
                        resultado = re.search(padrao, nome_arquivo)

                        # Verificar se encontrou o padrão
                        if resultado:
                            # Extrair o texto encontrado
                            extracao = resultado.group(1)
                        else:
                            print("Não foi possível encontrar o padrão.")

                        caminho_planilha = os.path.join(caminho_descarga_pendencia_recibo, 'CONTROLE_PENDENCIAS.xlsx')  # Ajuste para o caminho correto
                        df = pd.read_excel(caminho_planilha)

                        # Remover a linha onde a coluna 'LCD' tem o nome do arquivo
                        df = df[df["Nº LCD"] != extracao]

                        # Salvar a planilha sem a linha removida
                        df.to_excel(caminho_planilha, index=False)
                        print(f"✅ Linha com '{extracao}' removida da planilha!")
                        time.sleep(10)
                        break  # Para após encontrar a primeira correspondência
    else:                
        for pdfs in pdfs_lcd:
            
            padrao = r"LCD (.*?)\-PC"

            # Buscar a correspondência
            resultado = re.search(padrao, pdfs)

            # Verificar se encontrou o padrão
            if resultado:
                # Extrair o texto encontrado
                extracao = resultado.group(1)
            else:
                print("Não foi possível encontrar o padrão.")
            
            if num_recibo == extracao:
                print('vamos renomear')
                nome_novo_recibo = pdfs.replace("LCD ", "RECIBO LCD ").replace('NÃO ', '').replace('REEMBOLSÁVEL-', '')
                novo_local_recibo = os.path.join(caminho, nome_novo_recibo)
                os.rename(local_recibo, novo_local_recibo)
                shutil.move(novo_local_recibo, destino)

                arquivo_excel = os.path.join(caminho_descarga_lcd, 'PENDENCIAS_DESCARGAS.xlsx')

                nome_aba = "Sheet1"

                # Carregar a planilha
                df = pd.read_excel(arquivo_excel)

                # Novo valor que queremos atribuir na coluna "Status" quando encontrar o código
                novo_valor = "SIM"

                # Atualiza a coluna "Status" quando o valor for encontrado na coluna "Código"
                df.loc[df["Nº LCD"] == num_recibo, "POSSUI RECIBO?"] = novo_valor

                # Salva as alterações no arquivo Excel (no modo "w" para sobrescrever)
                df.to_excel(arquivo_excel, sheet_name=nome_aba, index=False)

                print("Alteração concluída!")

                break

# Caminho completo para o arquivo Excel
caminho_anexo = os.path.join(caminho_descarga_lcd, 'PENDENCIAS_DESCARGAS.xlsx')

# Carregar a planilha
Planilha_lcds = pd.read_excel(caminho_anexo)
qtd_pendencias = len(Planilha_lcds)

# Filtrar apenas as linhas onde "POSSUI RECIBO?" é "NAO"
Planilha_lcds_sem = Planilha_lcds[Planilha_lcds["POSSUI RECIBO?"].str.strip() == "NAO"]

# Caminho do arquivo de controle de pendências
caminho_controle_pendencias = os.path.join(caminho_descarga_pendencia_recibo, "CONTROLE_PENDENCIAS.xlsx")

# Verifica se a planilha de controle já existe
if os.path.exists(caminho_controle_pendencias):
    Planilha_controle = pd.read_excel(caminho_controle_pendencias)
else:
    Planilha_controle = pd.DataFrame()  # Criar um DataFrame vazio se não existir

# Adiciona as novas pendências ao histórico
Planilha_controle = pd.concat([Planilha_controle, Planilha_lcds_sem], ignore_index=True)
# Garantir que a coluna "PIX" seja numérica
Planilha_controle['PIX'] = pd.to_numeric(Planilha_controle['PIX'], errors='coerce')

# Remover duplicados baseado na coluna 'Nº LCD'
Planilha_controle = Planilha_controle.drop_duplicates(subset='Nº LCD', keep='last')

# Salvar o DataFrame atualizado no arquivo Excel com o openpyxl
with pd.ExcelWriter(caminho_controle_pendencias, engine='openpyxl') as writer:
    Planilha_controle.to_excel(writer, index=False)

    # Usando openpyxl para formatar a coluna "PIX" para evitar notação científica
    workbook = writer.book
    worksheet = workbook.active
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=6, max_col=6):  # Coluna PIX (7ª coluna)
        for cell in row:
            cell.number_format = '0'  # Formato numérico inteiro

print("Pendências adicionadas com sucesso!")

nome = f'GRADE {tempo_limite.strftime("%H-%M")} a {agora.strftime("%H-%M")}.xlsx'

try:
    print(os.path.join(caminho_descargas, nome))
    caminho_nova_planilha = os.path.join(caminho_descargas, nome)

    Planilha_grade = pd.read_excel(caminho_nova_planilha)

    # Pastas contendo os PDFs
    pastas_pdfs = [caminho, destino]  # Usa as variáveis já definidas

    if qtd_pendencias > 0:
        # Inicializar o corpo do email com o cabeçalho em formato HTML
        corpo_email = f"""
        <p>Resumo descargas:</p>
        <table border="1" cellpadding="5" cellspacing="0">
        <tr>
        """

        # Adicionar o cabeçalho da tabela ao corpo do email
        for col in Planilha_grade.columns:
            corpo_email += f"<th>{col}</th>"

        corpo_email += "</tr>"

        # Iterar sobre as linhas do DataFrame e adicionar ao corpo do email em formato HTML
        for index, row in Planilha_grade.iterrows():
            corpo_email += "<tr>"
            for item in row:
                corpo_email += f"<td>{item}</td>"
            corpo_email += "</tr>"

        corpo_email += "</table>"

        # Criar e-mail
        outlook = win32com.client.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = 'pendencias.financeiro@jettatransportes.com.br'
        mail.Subject = f'GRADE DE PAGTO LCD ({tempo_limite.strftime('%H:%M:%S')} as {agora.strftime('%H:%M:%S')})'
        mail.HTMLBody = corpo_email  # Usar HTMLBody para o corpo do email em HTML

        # Adicionar o arquivo Excel como anexo
        mail.Attachments.Add(caminho_nova_planilha)

        # Adicionar todos os PDFs das duas pastas como anexo
        for pasta in pastas_pdfs:
            caminho_pasta = Path(pasta)
            if caminho_pasta.exists():  # Verifica se a pasta existe
                for pdf in caminho_pasta.glob("*.pdf"):  # Lista todos os PDFs da pasta
                    mail.Attachments.Add(str(pdf))  # Adiciona o PDF ao email

        # Enviar o e-mail
        mail.Send()

        print("E-mail enviado com sucesso com o anexo planilha_envio.xlsx!")
except:
       # Inicializar o corpo do email com o cabeçalho em formato HTML
    corpo_email = f"""
        Nenhuma Descarga pendente para pagamento.
        
        <br><br><br><br>  <!-- Adiciona 4 quebras de linha -->

        Att. Robô
    """
    # Criar e-mail
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'pendencias.financeiro@jettatransportes.com.br'
    mail.Subject = f'GRADE DE PAGTO LCD ({tempo_limite.strftime('%H:%M:%S')} as {agora.strftime('%H:%M:%S')})'
    mail.HTMLBody = corpo_email  # Usar HTMLBody para o corpo do email em HTML
    # Enviar o e-mail
    mail.Send()
    print('Finalizando sem enviar nada')
    sys.exit()

# Caminhos para as subpastas ou arquivos dentro de DESCARGAS
caminho_lcd_pagar = os.path.join(caminho_descargas, "PASTA LCD A PAGAR")
caminho_pendencia = os.path.join(caminho_descargas, "PENDENCIA DE RECIBO")
caminho_recibo = os.path.join(caminho_descargas, "RECIBO DE PAGAMENTO")

# Função para mover arquivos de uma pasta de origem para a pasta de destino
def mover_arquivos(origem, destino):
    for item in os.listdir(origem):
        caminho_origem_item = os.path.join(origem, item)
        caminho_destino_item = os.path.join(destino, item)
        
        # Verifica se é um arquivo ou diretório
        if os.path.isfile(caminho_origem_item):
            shutil.move(caminho_origem_item, caminho_destino_item)
            print(f"Movido arquivo: {item} para {destino}")
        elif os.path.isdir(caminho_origem_item):
            shutil.move(caminho_origem_item, caminho_destino_item)
            print(f"Movido diretório: {item} para {destino}")

# Move os arquivos de cada pasta
mover_arquivos(caminho, caminho_lcd_pagar)
mover_arquivos(destino, caminho_recibo)

for i, linha in enumerate(Planilha_lcds_sem.index):
    lcd = Planilha_lcds_sem.loc[linha, "Nº LCD"]
    cod_parceiro = str(Planilha_lcds_sem.loc[linha, "Nº PARCEIRO COML."])
    cc = str(Planilha_lcds_sem.loc[linha, "CC"])
    reembolsavel = str(Planilha_lcds_sem.loc[linha, "REEMBOLSAVEL?"])
    pdf = 'LCD ' + lcd + '-PC ' + cod_parceiro +  '-'+ reembolsavel + '-CC' + cc
    pdf = os.path.join(caminho_lcd_pagar, f"{pdf}.pdf")

    #TRY DE TESTES
    try:
        # Verifica se o arquivo PDF existe antes de copiar
        if os.path.exists(caminho_pendencia):
            shutil.copy2(pdf, caminho_pendencia)
        else:
            print(f"Arquivo não encontrado: {caminho_pdf}")
    except:
        pass

# Lista todos os arquivos na pasta
arquivos = os.listdir(caminho_descarga_lcd)

# Percorre os arquivos
for arquivo in arquivos:
    if arquivo.endswith(".pdf") and not arquivo.startswith("LCD"):
        caminho_arquivo = os.path.join(caminho_descarga_lcd, arquivo)
        os.remove(caminho_arquivo)  # Remove o arquivo
        print(f"🗑 Arquivo removido: {arquivo}")

print("✅ Exclusão concluída.")

caminho_descarga_destino = os.path.join(caminho_descargas, 'GRADE') 

#Criar uma pasta com a data de hoje (formato YYYY-MM-DD)
data_hoje = datetime.now().strftime('%Y-%m-%d')
caminho_pasta_data = os.path.join(caminho_descarga_destino, data_hoje)

# Criar a pasta da data de hoje se não existir
if not os.path.exists(caminho_pasta_data):
    os.makedirs(caminho_pasta_data)
    print(f"Pasta criada: {caminho_pasta_data}")

caminho_nova_planilha = os.path.join(caminho_descargas, nome)

# Mover o arquivo para o destino
shutil.move(caminho_nova_planilha, caminho_pasta_data)

print("Processo concluído!")