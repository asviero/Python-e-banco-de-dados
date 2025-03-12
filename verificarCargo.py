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

def read_excel(input):
    df = pd.read_excel(input)
    if 'Matrícula (obrigatório)' not in df.columns or 'Nome do Segurado                                                                                  (obrigatório)' not in df.columns:
        print("A coluna 'matricula' ou a coluna 'nome' não foram encontradas no arquivo.")
        exit()
    return df

def search_cargo(matricula, nome, cursor):
    matricula = str(matricula)[:-1]

    cursor.execute("SELECT cs.MATR_ORIGEM_CONTRATO, cs.ID_CARGO, cgs.NOME_CARGO "
                    "FROM CONTRATO_SERV cs "
                    "JOIN CARGO_SERV cgs ON cs.ID_CARGO = cgs.ID_CARGO_SERV "
                    "WHERE cs.MATR_ORIGEM_CONTRATO = ?", (matricula,))
    resultado = cursor.fetchone()

    if resultado:
        return matricula, nome, resultado[2]
    else:
        return matricula, nome, "Não Encontrado"
    
def processar_dados(input, conn):
    cursor = conn.cursor()
    df = read_excel(input)

    log_data = []

    for _, row in df.iterrows():
        matricula = row['Matrícula (obrigatório)']
        nome = row['Nome do Segurado                                                                                  (obrigatório)']
        dados = search_cargo(matricula, nome, cursor)
        log_data.append({"matricula": dados[0], "nome": dados[1], "cargo": dados[2]})

    cursor.close()

    log_df = pd.DataFrame(log_data)
    log_df.to_excel("logCargos.xlsx", index=False)
    print("Arquivo 'logCargos.xlsx' criado com sucesso.")

if __name__ == "__main__":
    conn = connect_db()
    input = "C:/Users/angelo.viero/Documents/Projetos/Funcionarios/Aposentados/6-1. BENEFICIOS - APOSENTADOS - COMPLEMENTOS NECESSÁRIOS.xlsx"
    processar_dados(input, conn)
    conn.close()
    print("Programa finalizado com sucesso.")