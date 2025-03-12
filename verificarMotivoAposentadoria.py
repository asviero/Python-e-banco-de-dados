import fdb
import pandas as pd

motivos_dict = {
    '327.0': "não definido",
    '327.1': "Voluntária Serviço / Contribuição - Integral",
    '327.2': "Voluntária Serviço / Contribuição - Proporcional",
    '327.3': "Idade - Compulsória / Proporcional",
    '327.4': "Invalidez - Integral",
    '327.5': "Invalidez - Proporcional",
    '327.6': "Idade Voluntária / Proporcional"
}

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

def search_aposentadoria(matricula, nome, cursor):
    matricula = str(matricula)[:-1]

    cursor.execute("SELECT ITEM_MOTIVO_APOSENTADORIA "
                    "FROM CONTRATO_SERV "
                    "WHERE MATR_ORIGEM_CONTRATO = ?", (matricula,))
    resultado = cursor.fetchone()

    motivo = motivos_dict.get(resultado[0], "não definido") if resultado else "Não Encontrado"

    return matricula, nome, motivo

def processar_dados(input, conn):
    cursor = conn.cursor()
    df = read_excel(input)

    log_data = []

    for _, row in df.iterrows():
        matricula = row['Matrícula (obrigatório)']
        nome = row['Nome do Segurado                                                                                  (obrigatório)']
        dados = search_aposentadoria(matricula, nome, cursor)
        log_data.append({"matricula": dados[0], "nome": dados[1], "motivo": dados[2]})

    cursor.close()

    log_df = pd.DataFrame(log_data)
    log_df.to_excel("logMotivosAposentadoria.xlsx", index=False)
    print("Arquivo 'logMotivosAposentadoria.xlsx' criado com sucesso.")

if __name__ == '__main__':
    conn = connect_db()
    input = "C:/Users/angelo.viero/Documents/Projetos/Funcionarios/Aposentados/6-1. BENEFICIOS - APOSENTADOS - COMPLEMENTOS NECESSÁRIOS.xlsx"
    processar_dados(input, conn)
    conn.close()
    print("Programa finalizado com sucesso.")
