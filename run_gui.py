# ... importações ...
from sanitizer import DataSanitizer
from excel_handler import ExcelHandler

# Dentro da sua função de processamento (Worker ou Button Click):
def executar_processamento(caminho_arquivo_entrada, caminho_modelo):
    try:
        # 1. Instancia as classes
        sanitizer = DataSanitizer()
        handler = ExcelHandler(output_dir="Output") # ou a pasta que você usa
        
        # 2. Sanitização (Agora retorna DOIS valores)
        print("Lendo arquivo e metadados...")
        df_itens, info_header = sanitizer.sanitize(caminho_arquivo_entrada)
        
        if df_itens.empty:
            raise Exception("Nenhum item encontrado na tabela.")
            
        print(f"Itens encontrados: {len(df_itens)}")
        print(f"Dados do cabeçalho: {info_header}")
        
        # 3. Geração do Excel (Passando os DOIS valores)
        print("Gerando planilha final...")
        caminho_final = handler.processar_modelo_insert(
            modelo_path=caminho_modelo,
            df_dados=df_itens,
            info_cabecalho=info_header # <--- O PULO DO GATO ESTÁ AQUI
        )
        
        print(f"Concluído! Arquivo salvo em: {caminho_final}")
        return caminho_final

    except Exception as e:
        print(f"Erro Fatal: {e}")
        # Aqui você pode mostrar um popup de erro na sua GUI