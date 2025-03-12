import fdb
import pandas as pd

def load_config(configdb):
    config = {}
    with open(configdb, "r", encoding="utf-8") as file:
        for line in file:
            key, value = line.strip().split("=")
            config[key] = value
    return config

def connect_db():
    config = load_config("configdb.txt")

    conn = fdb.connect(
        host=config['host'],
        database=config['database'],
        user=config['user'],
        password=config['password']
    )
    return conn

def read_excel(file):
    df = pd.read_excel(file)
    if 'nome' not in df.columns:
        print("A coluna 'nome' não foi encontrada no arquivo.")
        exit()
    return df

def search_data_obito(nome, cursor):
    cursor.execute("SELECT NOME_PESSOA, DT_OBITO FROM PESSOA WHERE UPPER(NOME_PESSOA) = ?", [nome.strip().upper()])
    resultado = cursor.fetchone()

    if resultado and resultado[1]:
        return resultado
    return None

def processar_dados(input, conn):
    cursor = conn.cursor()
    df = read_excel(input)

    with open("logDt_Obito_TESTE.txt", "w", encoding="utf-8") as log_file:
        for nome in df['nome']:
            data_obito = search_data_obito(nome, cursor)
            if data_obito:
                log_entry = f"Nome: {data_obito[0]} | Data de Óbito: {data_obito[1]}\n"
                print(log_entry.strip())
                log_file.write(log_entry)
    
    cursor.close()

if __name__ == "__main__":
    conn = connect_db()
    input = "C:/Users/angelo.viero/Documents/Projetos/Funcionarios/Aposentados/aposentados_file.xlsx"
    processar_dados(input, conn)
    conn.close()
    print("Programa finalizado com sucesso.")
