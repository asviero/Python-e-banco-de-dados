import pandas as pd

def load_names_from_txt(txt_path):
    with open(txt_path, "r", encoding="utf-8") as f:
        return [line.strip().upper() for line in f.readlines()]

def load_excel(xlsx_path):
    return pd.read_excel(xlsx_path)

def remove_names_from_dataframe(df, names_list):
    initial_count = len(df)
    df = df[~df['nome'].isin(names_list)]
    deleted_count = initial_count - len(df)
    return df, deleted_count

def save_excel(df, xlsx_path):
    df.to_excel(xlsx_path, index=False)

def processar_dados(xlsx_path, txt_path):
    nomes_txt = load_names_from_txt(txt_path)
    df = load_excel(xlsx_path)
    df, deleted_count = remove_names_from_dataframe(df, nomes_txt)
    save_excel(df, xlsx_path)
    print(f"Total de nomes deletados: {deleted_count}")

if __name__ == "__main__":
    xlsx_path = "C:/Users/angelo.viero/Documents/Projetos/Funcionarios/Aposentados/aposentados_file.xlsx"
    txt_path = "C:/Users/angelo.viero/Documents/Projetos/Funcionarios/Aposentados/logDt_Óbito.txt"
    processar_dados(xlsx_path, txt_path)