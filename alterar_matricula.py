import pandas as pd

def limpar_matricula(matricula):
    matricula = str(matricula)

    if matricula[0] in ['2', '1']:
        matricula = matricula[1:]

    matricula = matricula.lstrip('0')
    matricula = matricula[:-2]
    return matricula

def processar_planilha(input, output):
    df = pd.read_excel(input)

    if 'matricula' in df.columns:
        df['matricula'] = df['matricula'].apply(limpar_matricula)

        df.to_excel(output, index=False)
        print(f"Arquivo salvo em: {output}")
    else:
        print("A coluna 'matricula' não foi encontrada na planilha.")

if __name__ == '__main__':
    input = "C:/Users/angelo.viero/Documents/Projetos/Funcionarios/Aposentados/aposentados_file.xlsx"
    output = "C:/Users/angelo.viero/Documents/Projetos/Funcionarios/Aposentados/aposentados_file_processado.xlsx"
    processar_planilha(input, output)
    print("Programa finalizado com sucesso.")