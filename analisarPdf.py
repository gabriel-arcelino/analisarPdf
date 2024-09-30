import pdfplumber
from openpyxl import Workbook

def extrair_dados_pdf(caminho_pdf):
    with pdfplumber.open(caminho_pdf) as pdf:
        nome_escola = None
        dados_turmas = {}
        turma_atual = None
        turma_identificacao_atual = []

        for pagina in pdf.pages:
            texto = pagina.extract_text()
            linhas = texto.split('\n')
            
            for i, linha in enumerate(linhas):
                if not nome_escola and 'ESCOLA:' in linha:
                    partes = linha.split(' - ')
                    if len(partes) > 1:
                        nome_escola = partes[1].strip()
                        print(f"Nome da escola encontrado: {nome_escola}")
                
                # Identificar a turma atual
                if "TURMA" in linha:
                    if i + 1 < len(linhas):  
                        id_turma = obter_id_turma(linhas,i)
                        turma_identificacao = linhas[i-1].strip().split("|")
                        print("tirma_id encontrada: ",turma_identificacao)
                        if not turma_atual:
                            turma_identificacao[-1] = id_turma
                            turma_atual = f"{' | '.join(turma_identificacao)}" 
                            dados_turmas[turma_atual] = 0  # Inicializar contagem para essa turma
                            turma_identificacao_atual = turma_identificacao
                            print(f"Turma inicial: {turma_atual}") 
                        elif turma_identificacao_atual[-1] != id_turma and not verificar_elementos_na_lista(['SEDUC'],linhas[i+1].split(" ")):
                            print("idturma_identificacao_atual ",turma_identificacao_atual[-1],"id_turma ",id_turma)
                            if verificar_elementos_na_lista(["Fundamental","Infantil"],turma_identificacao):
                                print("nova turma")
                                turma_identificacao_atual = turma_identificacao
                            else:
                                print("Mesmo ano")
                            turma_identificacao_atual[-1] = id_turma
                            turma_atual = f"{' | '.join(turma_identificacao_atual)}"
                            dados_turmas[turma_atual] = 0 
                            print(f"Turma atual definida: {turma_atual}")
                       
                        elif verificar_elementos_na_lista(["Fundamental","Infantil"],turma_identificacao):
                            if turma_identificacao_atual[0] != turma_identificacao[0]:                                  
                                turma_identificacao_atual = turma_identificacao
                                turma_identificacao_atual[-1] = id_turma
                                turma_atual = f"{' | '.join(turma_identificacao_atual)}"
                                dados_turmas[turma_atual] = 0 
                                print(f"Turma atual definida: {turma_atual}") 
                # Contar alunos
                elif turma_atual and linha.strip() and not verificar_elementos_na_lista(["SEDUC","Fortaleza","Fundamental","Infantil","NOMINAL","SECRETARIA","Alunos","ESCOLA","ANOS"],linha.split(" ")):
                    partes = linha.split()  
                    if len(partes) > 2: 
                        dados_turmas[turma_atual] += 1  
                        print("aluno encontrado: ",linha)
        return nome_escola, dados_turmas

def obter_id_turma(linha,i):  
    if verificar_elementos_na_lista(["CRECHE","ESCOLA"],linha[i+1].split(" ")) :
        return "Turma "+linha[i+1]+" "+linha[i+3].strip()
    return "Turma "+linha[i+1][0]
def verificar_elementos_na_lista(elementos,lista):
    for item in lista:  
        for elemento in elementos:
            if elemento in item:
                return True
    return False

def salvar_dados_excel(nome_escola, dados_turmas):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = nome_escola

    sheet.append(["Turma", "Quantidade de Alunos"])

    for turma, quantidade in dados_turmas.items():
        sheet.append([turma, quantidade])
        # print(f"Adicionando ao Excel: {turma} - {quantidade} alunos")

    # print(f"Salvando dados em {nome_escola}.xlsx")
    workbook.save(f"{nome_escola}.xlsx")
    print(f"Dados salvos em {nome_escola}.xlsx")

def buscar_linha_pdf(caminho_pdf, string_busca):
    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            linhas = texto.split('\n')
            for linha in linhas:
                if string_busca in linha:
                    return linha 
    return None
# Exemplo de uso
if __name__ == "__main__":
    caminho_pdf = "relatorio (1).pdf"
    nome_escola, dados_turmas = extrair_dados_pdf(caminho_pdf)
    if nome_escola and dados_turmas:
        salvar_dados_excel(nome_escola, dados_turmas)
    else:
        print("Não foi possível extrair os dados do PDF.")

