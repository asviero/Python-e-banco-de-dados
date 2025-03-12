import fdb
import pandas as pd

estado_civil_dict = {
    '11.1': "1",
    '11.2': "2",
    '11.3': "4",
    '11.4': "3",
    '11.5': "6",
    '11.6': "5"
}

sexo_dict = {'5.1': "1",
            '5.2': "2"}

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
    matricula = str(matricula)[:-1]
    cursor.execute("SELECT PESSOA.NOME_PESSOA, PESSOA.NOME_MAE, PESSOA.DT_NASCIMENTO, PESSOA.ITEM_ESTADO_CIVIL, PESSOA.ITEM_SEXO "
                    "FROM PESSOA "
                    "JOIN CONTRATO_SERV ON CONTRATO_SERV.ID_PESSOA = PESSOA.ID_PESSOA "
                    "WHERE CONTRATO_SERV.MATR_ORIGEM_CONTRATO = ?", (matricula,))
    resultado = cursor.fetchone()

    if resultado:
        nome_pessoa, nome_mae, dt_nascimento, estado_civil, sexo = resultado

        estado_civil = estado_civil_dict.get(estado_civil, "Outros")
        sexo = sexo_dict.get(sexo, sexo)

        return nome_pessoa, nome_mae, dt_nascimento, estado_civil, sexo
    else:
        return matricula, "Não Encontrado", "Não Encontrado", "Não Encontrado", "Não Encontrado"

def processar_dados(input, conn):
    cursor = conn.cursor()
    df = read_excel(input)

    log_data = []

    for matricula in df['Matrícula (obrigatório)']:
        dados = search_dados(matricula, cursor)
        log_data.append({"Matrícula": matricula, "Nome": dados[0], "Nome da Mãe": dados[1], "Data de Nascimento": dados[2], "Estado Civil": dados[3], "Sexo": dados[4]})

    cursor.close()

    log_df = pd.DataFrame(log_data)
    log_df.to_excel("logDados.xlsx", index=False)
    print("Arquivo 'logDados.xlsx' criado com sucesso.")

if __name__ == "__main__":
    conn = connect_db()
    input = "C:/Users/angelo.viero/Documents/Projetos/Funcionarios/Aposentados/6-1. BENEFICIOS - APOSENTADOS - COMPLEMENTOS NECESSÁRIOS.xlsx"
    processar_dados(input, conn)
    conn.close()
    print("Programa finalizado com sucesso.")