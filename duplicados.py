import pandas as pd
import os
from pathlib import Path

def processar_planilha(
    caminho_arquivo_entrada, 
    caminho_arquivo_saida=None,
    nome_coluna="NomeEmpresa",
    limite_frequencia=3,
    indice_coluna1=0,
    indice_coluna2=1
):
    """
    Processa uma planilha verificando igualdade entre colunas e frequência,
    salvando registros únicos em uma nova planilha.
    
    Args:
        caminho_arquivo_entrada (str): Caminho para o arquivo Excel de entrada
        caminho_arquivo_saida (str, opcional): Caminho para o arquivo de saída
        nome_coluna (str, opcional): Nome da coluna na planilha de saída
        limite_frequencia (int, opcional): Limite máximo de ocorrências permitidas
        indice_coluna1 (int, opcional): Índice da primeira coluna a ser verificada
        indice_coluna2 (int, opcional): Índice da segunda coluna a ser verificada
    
    Returns:
        str: Caminho do arquivo de saída criado
        dict: Estatísticas do processamento
    """
    try:
        # Verificar extensão do arquivo
        if not (caminho_arquivo_entrada.endswith('.xlsx') or 
                caminho_arquivo_entrada.endswith('.xls') or 
                caminho_arquivo_entrada.endswith('.csv')):
            raise ValueError("Formato de arquivo não suportado. Use xlsx, xls ou csv.")
        
        # Determinar o tipo de arquivo e carregar
        if caminho_arquivo_entrada.endswith('.csv'):
            df = pd.read_csv(caminho_arquivo_entrada)
        else:
            df = pd.read_excel(caminho_arquivo_entrada)
        
        # Verificar se existem pelo menos duas colunas
        if len(df.columns) <= max(indice_coluna1, indice_coluna2):
            raise ValueError(f"A planilha deve conter pelo menos {max(indice_coluna1, indice_coluna2) + 1} colunas.")
        
        # Obter os nomes reais das colunas pelos índices
        coluna1 = df.columns[indice_coluna1]
        coluna2 = df.columns[indice_coluna2]
        
        # Verificar se as colunas são iguais
        colunas_iguais = df[coluna1].equals(df[coluna2])
        
        # Estatísticas para retornar
        estatisticas = {
            "colunas_iguais": colunas_iguais,
            "valores_acima_limite": [],
            "total_registros_entrada": len(df),
            "total_registros_saida": 0
        }
        
        # Contando ocorrências (considerando ambas as colunas)
        valores_combinados = pd.concat([df[coluna1], df[coluna2]])
        contagem = valores_combinados.value_counts()
        
        # Identificando valores que aparecem mais que o limite
        valores_frequentes = contagem[contagem > limite_frequencia]
        estatisticas["valores_acima_limite"] = valores_frequentes.to_dict()
        
        # Criando o dataframe de saída com valores únicos
        valores_unicos = valores_combinados.drop_duplicates().dropna()
        df_saida = pd.DataFrame({nome_coluna: valores_unicos})
        estatisticas["total_registros_saida"] = len(df_saida)
        
        # Definindo caminho de saída, se não foi especificado
        if not caminho_arquivo_saida:
            diretorio = os.path.dirname(caminho_arquivo_entrada)
            nome_base = Path(caminho_arquivo_entrada).stem
            caminho_arquivo_saida = os.path.join(diretorio, f"{nome_base}_processado.xlsx")
        
        # Salvar o resultado
        df_saida.to_excel(caminho_arquivo_saida, index=False)
        
        return caminho_arquivo_saida, estatisticas
    
    except Exception as e:
        print(f"Erro ao processar o arquivo: {str(e)}")
        return None, {"erro": str(e)}

# Exemplo de uso com script independente
if __name__ == "__main__":
    import sys
    
    # Configurações padrão que podem ser modificadas diretamente no código
    CONFIG = {
        "caminho_entrada": "./contatos.xlsx",
        "caminho_saida": None,
        "nome_coluna": "NomeEmpresa",
        "limite_frequencia": 3,
        "indice_coluna1": 0,
        "indice_coluna2": 1
    }
    
    # Verificar argumentos de linha de comando
    if len(sys.argv) > 1:
        CONFIG["caminho_entrada"] = sys.argv[1]
    if len(sys.argv) > 2:
        CONFIG["caminho_saida"] = sys.argv[2]
    
    # Se não tiver caminho de entrada e estiver rodando como script principal, solicitar
    if not CONFIG["caminho_entrada"]:
        CONFIG["caminho_entrada"] = input("Digite o caminho do arquivo de entrada: ")
    
    # Execução principal
    caminho_saida, estatisticas = processar_planilha(
        CONFIG["caminho_entrada"],
        CONFIG["caminho_saida"],
        CONFIG["nome_coluna"],
        CONFIG["limite_frequencia"],
        CONFIG["indice_coluna1"],
        CONFIG["indice_coluna2"]
    )
    
    # Exibir resultados
    if caminho_saida:
        print(f"\nArquivo processado com sucesso e salvo em: {caminho_saida}")
        print(f"Total de registros na entrada: {estatisticas['total_registros_entrada']}")
        print(f"Total de registros únicos na saída: {estatisticas['total_registros_saida']}")
        
        if not estatisticas["colunas_iguais"]:
            print("Aviso: As duas colunas não são idênticas.")
            
        if estatisticas["valores_acima_limite"]:
            print(f"Aviso: Os seguintes valores aparecem mais de {CONFIG['limite_frequencia']} vezes:")
            for valor, contagem in estatisticas["valores_acima_limite"].items():
                print(f"  - '{valor}': {contagem} vezes")
