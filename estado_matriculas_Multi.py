import requests
from bs4 import BeautifulSoup
import pandas as pd
import threading
import logging
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

# Configurar o logger para salvar as mensagens no arquivo log.txt
logging.basicConfig(filename='log.txt', level=logging.INFO, format='%(asctime)s - %(message)s')

# Função para converter matrícula para o formato 'xx-xx-xx'
def convert_matricula_format(matricula):
    if '-' not in matricula:
        return f"{matricula[:2]}-{matricula[2:4]}-{matricula[4:]}"
    return matricula

# Função para verificar uma única matrícula com repetição em caso de falha
def check_matricula(matricula, retries=3, delay=5):
    url = 'https://www.imt-ip.pt/matriculascanceladas/matriculas.asp'
    data = {'matricula': matricula}
    for attempt in range(retries):
        try:
            response = requests.post(url, data=data)
            response.raise_for_status()  # Levanta um erro para status de resposta HTTP ruim
            soup = BeautifulSoup(response.content, 'html.parser')
            mensagem = soup.find('div', class_='mensagem').text.strip()
            return mensagem
        except requests.RequestException as e:
            logging.error(f"Tentativa {attempt + 1} falhou para a matrícula {matricula}: {e}")
            time.sleep(delay)
    raise Exception(f"Falha ao analisar a matrícula {matricula} após {retries} tentativas")

# Função para processar uma linha do DataFrame
def process_row(index, row, mensagens, status, total_matriculas):
    try:
        matricula = row['Matricula']
        matricula = convert_matricula_format(matricula)
        mensagem = check_matricula(matricula)
        mensagens[index] = mensagem
        if 'não' in mensagem:
            status[index] = 'Não Cancelada'
        else:
            status[index] = 'Cancelada'
        log_message = f"Analisadas {index + 1} de {total_matriculas} matrículas"
        print(log_message)
        logging.info(log_message)
    except Exception as e:
        error_message = f"Erro ao analisar a matrícula {matricula}: {e}"
        print(error_message)
        logging.error(error_message)
        raise

# Carregar o arquivo Excel
excel_file = 'matriculas.xlsx'
df = pd.read_excel(excel_file, engine='openpyxl')

# Listas para armazenar resultados
mensagens = [None] * len(df)
status = [None] * len(df)

# Total de matrículas
total_matriculas = len(df)

# Usar ThreadPoolExecutor para gerenciar threads
with ThreadPoolExecutor(max_workers=10) as executor:
    futures = [executor.submit(process_row, index, row, mensagens, status, total_matriculas) for index, row in df.iterrows()]
    for future in as_completed(futures):
        try:
            future.result()
        except Exception as e:
            print(f"Erro em uma das threads: {e}")
            logging.error(f"Erro em uma das threads: {e}")

# Adicionar as mensagens e os status a novas colunas no DataFrame
df['Mensagem'] = mensagens
df['Status'] = status

# Salvar o DataFrame atualizado em um novo arquivo Excel
df.to_excel('matriculas_resultados_2.xlsx', index=False)

success_message = "Script executado com sucesso. Os resultados foram salvos em 'matriculas_resultados.xlsx'."
print(success_message)
logging.info(success_message)
