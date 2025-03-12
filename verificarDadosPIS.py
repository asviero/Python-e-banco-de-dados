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
    if 'Matrícula (obrigatório)' not in df.columns:
        print("A coluna 'Matrícula (obrigatório)' não foi encontrada no arquivo.")
        exit()
    return df

def search_dados(matricula, cursor):
    matricula = str(matricula).strip()[:-1]
    cursor.execute("""
        SELECT p.NOME_PESSOA, p.ID_PESSOA, p.DOC_PIS, cs.MATR_ORIGEM_CONTRATO
        FROM PESSOA p
        INNER JOIN CONTRATO_SERV cs ON cs.ID_PESSOA = p.ID_PESSOA
        WHERE cs.MATR_ORIGEM_CONTRATO = ?
    """, (matricula,))
    resultado = cursor.fetchone()

    if resultado:
        return resultado[0], resultado[3], resultado[2]
    else:
        return matricula, "Não encontrado", "Não encontrado"

def processar_dados(input, conn):
    cursor = conn.cursor()
    df = read_excel(input)

    log_data = []

    for matricula in df['Matrícula (obrigatório)']:
        pis = search_dados(matricula, cursor)
        log_data.append({"Nome": pis[0], "Matrícula": pis[1], "PASEP/PIS/NIT": pis[2]})

    cursor.close()

    log_df = pd.DataFrame(log_data)
    log_df.to_excel("logPIS.xlsx", index=False)
    print("Arquivo 'logPIS.xlsx' criado com sucesso.")

if __name__ == "__main__":
    conn = connect_db()
    input = "C:/Users/angelo.viero/Documents/Projetos/Funcionarios/Aposentados/6-1. BENEFICIOS - APOSENTADOS - COMPLEMENTOS NECESSÁRIOS.xlsx"
    processar_dados(input, conn)
    conn.close()
    print("Programa finalizado com sucesso.")