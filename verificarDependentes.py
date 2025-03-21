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

    required_columns = ['Matrícula (obrigatório)', 'Nome do Segurado                                                                                  (obrigatório)']

    for col in required_columns:
        if col not in df.columns:
            print(f"A coluna '{col}' não foi encontrada no arquivo.")
            exit()
    return df

def search_dados(matricula, cursor):
    matricula = str(matricula).strip()[:-1]
    cursor.execute("""
        SELECT
            (SELECT p2.NOME_PESSOA FROM PESSOA p2 WHERE p2.ID_PESSOA = d.ID_PESSOA_TITULAR) AS "titular aposentado",
            p.NOME_PESSOA AS dependentes,
            p.DOC_CPF AS "cpf dependente",
            (SELECT ITA2.DESCR_ITEM_AUX FROM ITEM_TAB_AUXILIAR ita2 WHERE ITA2.COD_ESTRUT_ITEM= d.ITEM_GRAU_DEPENDENCIA) AS "grau de dep",
            d.IND_SITUACAO AS "Sit. dependencia", (select p1.NOME_PESSOA from PESSOA p1 where p1.ID_PESSOA = d.ID_PESSOA_TITULAR) AS "nome"
        FROM DEPENDENTE d
        INNER JOIN PESSOA p ON d.ID_PESSOA = p.ID_PESSOA
        WHERE d.ID_PESSOA_TITULAR IN (
            SELECT cs.ID_PESSOA
            FROM CONTRATO_SERV cs
            WHERE cs.ITEM_SITUACAO_FUNCIONAL = '331.6'
            AND cs.IND_SITUACAO = 'A'
            AND cs.MATR_ORIGEM_CONTRATO = ?
        )
        ORDER BY "titular aposentado" ASC
    """, (matricula,))

    resultado = cursor.fetchall()

    return resultado

def processar_dados(input, conn):
    cursor = conn.cursor()
    df = read_excel(input)

    log_data = []

    for matricula, nome in zip(df['Matrícula (obrigatório)'], df['Nome do Segurado                                                                                  (obrigatório)']):
        dependentes = search_dados(matricula, cursor)

        if dependentes:
            for dep in dependentes:
                log_data.append({
                    "Nome do Aposentado": nome,
                    "Matrícula": matricula,
                    "Nome do Dependente": dep[1],
                    "CPF do Dependente": dep[2],
                    "Grau Dependência": dep[3],
                    "Situação Dependência": dep[4]
                })
        else:
            log_data.append({
                "Nome do Aposentado": nome,
                "Matrícula": matricula,
                "Nome do Dependente": "Não encontrado",
                "CPF do Dependente": "Não encontrado",
                "Grau Dependência": "Não encontrado",
                "Situação Dependência": "Não encontrado"
            })

    cursor.close()

    log_df = pd.DataFrame(log_data)
    log_df.to_excel("logDependentes.xlsx", index=False)
    print("Arquivo 'logDependentes.xlsx' criado com sucesso.")

if __name__ == "__main__":
    conn = connect_db()
    input = "C:/Users/angelo.viero/Documents/Projetos/Funcionarios/Aposentados/6-1. BENEFICIOS - APOSENTADOS - COMPLEMENTOS NECESSÁRIOS.xlsx"
    processar_dados(input, conn)
    conn.close()
    print("Programa finalizado com sucesso.")
