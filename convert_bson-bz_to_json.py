import os
import gzip
import subprocess
import pandas as pd

# Caminho para a pasta com os arquivos .bson.gz e .json.gz
PASTA_ORIGEM = r'caminho_para_a_pasta_onde_estão_os_arquivos_.bson.gz_e_.json.gz'
CAMINHO_BSONDUMP = r'caminho_para_a_pasta_com_mongodb-database-tools'

# Subpastas de saída
PASTA_SAIDA_JSON = os.path.join(PASTA_ORIGEM, 'json_convertidos')
PASTA_SAIDA_BSON = os.path.join(PASTA_ORIGEM, 'bson_descompactados')
PASTA_SAIDA_XLSX = os.path.join(PASTA_ORIGEM, 'xlsx_convertidos')

# Cria as pastas, se não existirem
os.makedirs(PASTA_SAIDA_JSON, exist_ok=True)
os.makedirs(PASTA_SAIDA_BSON, exist_ok=True)
os.makedirs(PASTA_SAIDA_XLSX, exist_ok=True)

# Descompacta e converte arquivos .bson.gz e .json.gz
for nome_arquivo in os.listdir(PASTA_ORIGEM):
    caminho_arquivo = os.path.join(PASTA_ORIGEM, nome_arquivo)

    if nome_arquivo.endswith('.bson.gz'):
        nome_base_com_bson = os.path.splitext(nome_arquivo)[0]  # Remove .gz
        nome_base = os.path.splitext(nome_base_com_bson)[0]     # Remove .bson

        caminho_bson = os.path.join(PASTA_SAIDA_BSON, f'{nome_base}.bson')
        caminho_json = os.path.join(PASTA_SAIDA_JSON, f'{nome_base}.json')

        print(f'Descompactando BSON: {nome_arquivo} → {caminho_bson}')
        with gzip.open(caminho_arquivo, 'rb') as f_in:
            with open(caminho_bson, 'wb') as f_out:
                f_out.write(f_in.read())

        print(f'Convertendo BSON para JSON: {caminho_json}')
        comando = [CAMINHO_BSONDUMP, caminho_bson, '--outFile', caminho_json]

        try:
            subprocess.run(comando, check=True)
            print(f'{nome_base}.json criado com sucesso!')
        except subprocess.CalledProcessError as e:
            print(f'Erro ao converter {nome_base}: {e}')

    elif nome_arquivo.endswith('.json.gz'):
        nome_base = os.path.splitext(nome_arquivo)[0]  # Remove .gz
        caminho_json = os.path.join(PASTA_SAIDA_JSON, nome_base)

        print(f'Descompactando JSON: {nome_arquivo} → {caminho_json}')
        with gzip.open(caminho_arquivo, 'rb') as f_in:
            with open(caminho_json, 'wb') as f_out:
                f_out.write(f_in.read())

        print(f'{nome_base} descompactado com sucesso!')

# Converter arquivos JSON para XLSX
print("\nIniciando conversão de JSON para XLSX...\n")
for nome_arquivo_json in os.listdir(PASTA_SAIDA_JSON):
    if nome_arquivo_json.endswith('.json'):
        caminho_json = os.path.join(PASTA_SAIDA_JSON, nome_arquivo_json)
        nome_base = os.path.splitext(nome_arquivo_json)[0]
        caminho_xlsx = os.path.join(PASTA_SAIDA_XLSX, f'{nome_base}.xlsx')

        try:
            # Tenta carregar como JSON lines 
            print(f'Convertendo para XLSX...')
            with open(caminho_json, 'r', encoding='utf-8') as f:
                primeira_linha = f.readline()
                f.seek(0)
                if primeira_linha.strip().startswith('{'):
                    df = pd.read_json(f, lines=True)
                else:
                    df = pd.read_json(f)

            df.to_excel(caminho_xlsx, index=False, engine='openpyxl')
            print(f'Convertido para XLSX: {caminho_xlsx}')
        except Exception as e:
            print(f'Erro ao converter {nome_arquivo_json} para XLSX: {e}')

print('\nProcesso completo! Todos os arquivos foram processados.')
