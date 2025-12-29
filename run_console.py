"""
SISORC ULTIMATE - Modo Console (Blindado)
Use este script para ignorar erros gráficos e gerar o orçamento via texto.
"""
import os
import sys
from datetime import datetime
import pandas as pd

# Garante que acha os módulos
sys.path.append(os.getcwd())

from core.sanitizer import ExcelSanitizer
from core.excel_handler import OrcamentoEngine
from core.database import DatabaseManager
from utils.logger import Logger
from utils.helpers import ConfigLoader

def main():
    print("\n" + "="*60)
    print("🚀 SISORC ULTIMATE - MODO CONSOLE (SEM JANELAS)")
    print("="*60)

    # 1. Carregar Config
    try:
        config = ConfigLoader.carregar('config/settings.json')
        print("✅ Configurações carregadas.")
    except Exception as e:
        print(f"❌ Erro crítico na config: {e}")
        return

    # 2. Iniciar Componentes
    logger = Logger(nome="SISORC_CONSOLE")

    sanitizer = ExcelSanitizer(config.get('sanitizacao', {}))
    engine = OrcamentoEngine(config)
    db = DatabaseManager(config)

    # 3. Listar Arquivos Excel na pasta
    print("\n📂 Arquivos disponíveis nesta pasta:")
    arquivos = [f for f in os.listdir('.') if f.endswith('.xlsx') and not f.startswith('~$') and not f.startswith('temp_')]

    if not arquivos:
        print("❌ Nenhum arquivo .xlsx encontrado na pasta do script!")
        return

    for i, arq in enumerate(arquivos):
        print(f"   [{i+1}] {arq}")

    # 4. Seleção Manual
    try:
        idx_sint = int(input("\n👉 Digite o NÚMERO do arquivo SINTÉTICO (Dados): ")) - 1
        path_sintetico = arquivos[idx_sint]
        print(f"   Selecionado: {path_sintetico}")

        idx_mod = int(input("\n👉 Digite o NÚMERO do arquivo MODELO (Template): ")) - 1
        path_modelo = arquivos[idx_mod]
        print(f"   Selecionado: {path_modelo}")

        # Parâmetros adicionais
        print("\n⚙️ Parâmetros:")
        linha_ini = int(input("   Linha Inicial dos Dados (ex: 25): ") or 25)
        qtd_linhas = int(input("   Quantidade de Linhas a ler (ex: 100): ") or 100)

    except (ValueError, IndexError):
        print("❌ Opção inválida. Tente novamente.")
        return

    # 5. Dados do Projeto
    print("\n📝 Dados do Projeto:")
    obra = input("   Nome da Obra: ").strip() or "OBRA SEM NOME"
    local = input("   Local: ").strip() or "LOCAL INDEFINIDO"
    bdi_str = input("   BDI (ex: 25.5): ").replace(',', '.')

    try:
        bdi = float(bdi_str)
    except:
        bdi = 0.0
        print("   ⚠️ BDI inválido, assumindo 0.0%")

    dados_projeto = {'obra': obra, 'local': local, 'bdi': bdi}

    # 6. Execução
    print("\n⏳ Processando... (Aguarde)")

    # Passo A: Sanitizar
    print("   1/3 Sanitizando arquivo...")
    sucesso_san, arq_limpo, linha_header = sanitizer.sanitizar_arquivo(path_sintetico)

    if not sucesso_san:
        print(f"❌ Erro na sanitização: {arq_limpo}")
        sanitizer.limpar_arquivos_temp()
        return

    # Passo B: Preparar Níveis (Mockup para console)
    print("   2/3 Analisando dados...")
    mapa_niveis = {}
    try:
        # Lê rapidamente para tentar inferir níveis (simples)
        skip = linha_ini - 1
        df_preview = pd.read_excel(arq_limpo, header=None, skiprows=skip, nrows=qtd_linhas)
        # Assume coluna 0 como item se não mapeado, ou deixa o engine inferir
        # O engine tem lógica de inferência se o mapa retornar None/Vazio.
        pass
    except Exception as e:
        print(f"Aviso: Não foi possível pré-analisar níveis: {e}")

    # Passo C: Processar
    print("   3/3 Gerando orçamento...")
    intervalo = (linha_ini, linha_ini + qtd_linhas - 1)

    sucesso, resultado, stats = engine.processar_orcamento(
        arq_limpo,
        path_modelo,
        dados_projeto,
        mapa_niveis,
        intervalo
    )

    # Passo D: Limpeza
    sanitizer.limpar_arquivos_temp()

    if sucesso:
        print("\n" + "="*60)
        print("✅ SUCESSO! ORÇAMENTO GERADO.")
        print(f"📂 Arquivo: {resultado}")
        print("="*60)

        # Salva no banco
        try:
            db.inserir_orcamento({
                'data_geracao': datetime.now().isoformat(),
                'nome_obra': obra,
                'local': local,
                'bdi': bdi,
                'valor_total': 0, # Stats não retorna valor total ainda
                'arquivo_saida': resultado,
                'num_itens': stats.get('itens', 0),
                'num_titulos': 0,
                'duracao_processamento': 0
            })
        except: pass

        try:
            os.startfile(resultado)
        except:
            pass
    else:
        print(f"\n❌ ERRO NO ENGINE: {resultado}")

    input("\nPressione ENTER para fechar...")

if __name__ == "__main__":
    main()