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
    if 'id_contrato' not in df.columns:
        print("A coluna 'id_contrato' não foi encontrada no arquivo.")
        exit()
    return df

def count_lines(id_contrato, cursor):
    cursor.execute("""
        SELECT COUNT(*)
        FROM V_CONTRIB_PREVIDENCIA
        WHERE ID_CONTRATO_SERV = ?
    """, (id_contrato,))
    resultado = cursor.fetchone()

    if resultado:
        return id_contrato, resultado[0]
    else:
        return id_contrato, 0
    
def process_data(input, conn):
    cursor = conn.cursor()
    df = read_excel(input)

    log_data = []

    for id_contrato in df['id_contrato']:
        dados = count_lines(id_contrato, cursor)
        log_data.append({"id_contrato": dados[0], "linhas": dados[1]})

    cursor.close()

    log_df = pd.DataFrame(log_data)
    log_df.to_excel("logTotalLinesPerIdContrato.xlsx", index=False)
    print("Arquivo 'logtotalLinesPerIdContrato.xlsx' criado com sucesso.")

if __name__ == "__main__":
    conn = connect_db()
    input_file = "input.xlsx"
    process_data(input_file, conn)
    conn.close()
    print("Programa finalizado com sucesso.")