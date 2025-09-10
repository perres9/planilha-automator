import pandas as pd
import os
from datetime import datetime

def processar_planilhas():
    print("Iniciando processamento de planilhas...")
    
    # Criar pastas se não existirem
    os.makedirs('planilhas_entrada', exist_ok=True)
    os.makedirs('resultados', exist_ok=True)
    os.makedirs('processadas', exist_ok=True)
    
    # Procurar planilhas
    planilhas = [f for f in os.listdir('planilhas_entrada') 
                if f.endswith(('.xlsx', '.xls'))]
    
    if not planilhas:
        print("Nenhuma planilha encontrada na pasta 'planilhas_entrada'")
        return
    
    dados_combinados = []
    
    for planilha in planilhas:
        try:
            caminho = os.path.join('planilhas_entrada', planilha)
            print(f"Processando: {planilha}")
            
            # Ler planilha
            df = pd.read_excel(caminho)
            df['Arquivo_Origem'] = planilha
            df['Data_Processamento'] = datetime.now()
            
            dados_combinados.append(df)
            
            # Mover para processadas
            os.rename(caminho, os.path.join('processadas', planilha))
            
        except Exception as e:
            print(f"Erro em {planilha}: {e}")
    
    if dados_combinados:
        # Juntar todos os dados
        resultado = pd.concat(dados_combinados, ignore_index=True)
        
        # Salvar resultados
        data_hora = datetime.now().strftime("%Y%m%d_%H%M%S")
        arquivo_saida = f'resultados/consolidado_{data_hora}.xlsx'
        
        resultado.to_excel(arquivo_saida, index=False)
        print(f"Processamento concluído! Salvo em: {arquivo_saida}")
    else:
        print("Nenhum dado foi processado.")

if __name__ == "__main__":
    processar_planilhas()
