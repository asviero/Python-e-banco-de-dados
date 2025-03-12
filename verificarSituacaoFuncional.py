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

def verif_situacao_funcional(nome, cursor):
    cursor.execute("SELECT ID_PESSOA FROM PESSOA WHERE NOME_PESSOA = ?", (nome,))
    resultado = cursor.fetchone()

    if resultado:
        id_pessoa = resultado[0]
        cursor.execute("SELECT ITEM_SITUACAO_FUNCIONAL FROM CONTRATO_SERV WHERE ID_PESSOA = ?", (id_pessoa,))
        resultado_contrato = cursor.fetchone()

        if resultado_contrato:
            return resultado_contrato[0]
    return None

def processar_dados(arquivo_xlsx, conn):
    cursor = conn.cursor()
    df = read_excel(arquivo_xlsx)
#Alterar o nome da pasta de acordo com a Situação Funcional
    with open("logObito.txt", "w", encoding="utf-8") as log_file:
        for nome in df['nome']:
            situacao_funcional = verif_situacao_funcional(nome, cursor)
            if situacao_funcional is not None:
                try:
                    situacao_funcional = float(situacao_funcional)
                    if situacao_funcional == 331.8:
                        log_entry = f"Nome: {nome} | Código de Situação: {situacao_funcional}\n"
                        print(log_entry.strip())
                        log_file.write(log_entry)
                except ValueError:
                    print(f"Erro ao converter situação funcional para float: {situacao_funcional}")
    cursor.close()

#331.1: Em Atividade
#331.6: Aposentado
#331.7: Exonerado ( Rescisão Contrato)
#331.8: Óbito - Servidor em Atividade

if __name__ == "__main__":
    conn = connect_db()
    arquivo_xlsx = "C:/Users/angelo.viero/Documents/Projetos/Funcionarios/Aposentados/aposentados_file.xlsx"
    processar_dados(arquivo_xlsx, conn)
    conn.close()
    print("Programa finalizado com sucesso.")