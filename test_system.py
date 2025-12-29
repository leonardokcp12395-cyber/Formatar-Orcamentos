"""
SISORC ULTIMATE - Script de Teste e Diagnóstico
Use este script para validar se tudo está funcionando
"""

import sys
import os
from pathlib import Path

def print_section(title):
    """Imprime seção formatada"""
    print("\n" + "="*60)
    print(f"  {title}")
    print("="*60)

def test_imports():
    """Testa se todos os módulos podem ser importados"""
    print_section("1. TESTANDO IMPORTAÇÕES")
    
    modules = [
        ('pandas', 'Pandas'),
        ('openpyxl', 'OpenPyXL'),
        ('customtkinter', 'CustomTkinter'),
        ('PIL', 'Pillow')
    ]
    
    success = True
    for module_name, display_name in modules:
        try:
            __import__(module_name)
            print(f"  ✓ {display_name} OK")
        except ImportError as e:
            print(f"  ✗ {display_name} FALTANDO - {e}")
            success = False
    
    return success

def test_structure():
    """Testa se a estrutura de pastas está correta"""
    print_section("2. TESTANDO ESTRUTURA")
    
    required = {
        'Pastas': ['core', 'ui', 'utils', 'config'],
        'Core': ['core/sanitizer.py', 'core/excel_handler.py', 'core/database.py'],
        'UI': ['ui/main_window.py'],
        'Utils': ['utils/logger.py', 'utils/helpers.py'],
        'Config': ['config/settings.json'],
        'Scripts': ['main.py', 'run_gui.py', 'run_console.py']
    }
    
    success = True
    for category, items in required.items():
        print(f"\n  {category}:")
        for item in items:
            path = Path(item)
            if path.exists():
                print(f"    ✓ {item}")
            else:
                print(f"    ✗ {item} FALTANDO")
                success = False
    
    return success

def test_config():
    """Testa se a configuração está válida"""
    print_section("3. TESTANDO CONFIGURAÇÃO")
    
    try:
        from utils.helpers import ConfigLoader
        config = ConfigLoader.carregar('config/settings.json')
        
        required_keys = ['sanitizacao', 'excel']
        
        for key in required_keys:
            if key in config:
                print(f"  ✓ Seção '{key}' OK")
            else:
                print(f"  ✗ Seção '{key}' FALTANDO")
                return False
        
        return True
        
    except Exception as e:
        print(f"  ✗ Erro ao carregar config: {e}")
        return False

def test_logger():
    """Testa o sistema de log"""
    print_section("4. TESTANDO LOGGER")
    
    try:
        from utils.logger import Logger
        
        logger = Logger("TESTE")
        
        # Testa cada nível
        Logger.info("Teste de INFO")
        Logger.warning("Teste de WARNING")
        Logger.debug("Teste de DEBUG")
        Logger.error("Teste de ERROR")
        
        print("  ✓ Logger funcional")
        return True
        
    except Exception as e:
        print(f"  ✗ Erro no logger: {e}")
        return False

def test_sanitizer():
    """Testa o sanitizador"""
    print_section("5. TESTANDO SANITIZER")
    
    try:
        from core.sanitizer import ExcelSanitizer
        
        config = {'timeout': 30, 'limpar_temp': True}
        sanitizer = ExcelSanitizer(config)
        
        print("  ✓ Sanitizer instanciado")
        
        # Lista arquivos Excel na pasta atual
        excel_files = list(Path('.').glob('*.xlsx'))
        
        if excel_files:
            print(f"  ℹ️  Encontrados {len(excel_files)} arquivos .xlsx na pasta")
            print("     Para testar sanitização completa, execute o programa principal")
        else:
            print("  ⚠️  Nenhum arquivo .xlsx encontrado para teste")
        
        return True
        
    except Exception as e:
        print(f"  ✗ Erro no sanitizer: {e}")
        return False

def test_engine():
    """Testa o engine de orçamento"""
    print_section("6. TESTANDO ENGINE")
    
    try:
        from core.excel_handler import OrcamentoEngine
        
        config = {'excel': {'linha_inicial_modelo': 25}}
        engine = OrcamentoEngine(config)
        
        print("  ✓ Engine instanciado")
        print("  ℹ️  Para testar processamento completo, use o programa principal")
        
        return True
        
    except Exception as e:
        print(f"  ✗ Erro no engine: {e}")
        return False

def test_database():
    """Testa o banco de dados"""
    print_section("7. TESTANDO DATABASE")
    
    try:
        from core.database import DatabaseManager
        
        config = {'database': {'nome_arquivo': 'test_sisorc.db'}}
        db = DatabaseManager(config)
        
        print("  ✓ Database inicializado")
        
        # Testa estatísticas
        stats = db.buscar_estatisticas()
        print(f"  ✓ Estatísticas obtidas: {stats['total_orcamentos']} orçamentos")
        
        # Remove arquivo de teste
        try:
            Path('test_sisorc.db').unlink()
            print("  ✓ Arquivo de teste removido")
        except:
            pass
        
        return True
        
    except Exception as e:
        print(f"  ✗ Erro no database: {e}")
        return False

def test_ui():
    """Testa se a interface pode ser carregada"""
    print_section("8. TESTANDO INTERFACE")
    
    try:
        # Tenta importar sem inicializar
        from ui.main_window import SisorcApp
        
        print("  ✓ Interface pode ser importada")
        print("  ℹ️  Execute 'python run_gui.py' para abrir a interface")
        
        return True
        
    except Exception as e:
        print(f"  ✗ Erro na interface: {e}")
        return False

def create_example_files():
    """Cria arquivos de exemplo se não existirem"""
    print_section("9. ARQUIVOS DE EXEMPLO")
    
    # Verifica se há arquivos Excel
    excel_files = list(Path('.').glob('*.xlsx'))
    
    if not excel_files:
        print("  ⚠️  Nenhum arquivo Excel encontrado")
        print("  💡 Dica: Coloque seus arquivos SINTÉTICO e MODELO na pasta do programa")
    else:
        print("  ✓ Arquivos Excel encontrados:")
        for f in excel_files:
            size_mb = f.stat().st_size / (1024*1024)
            print(f"     - {f.name} ({size_mb:.2f} MB)")

def print_summary(results):
    """Imprime resumo dos testes"""
    print_section("RESUMO DOS TESTES")
    
    total = len(results)
    passed = sum(results.values())
    failed = total - passed
    
    print(f"\n  Total de testes: {total}")
    print(f"  ✓ Aprovados: {passed}")
    print(f"  ✗ Falharam: {failed}")
    
    if failed == 0:
        print("\n  🎉 TODOS OS TESTES PASSARAM!")
        print("  🚀 Sistema pronto para uso")
        print("\n  Execute: python run_gui.py")
    else:
        print("\n  ⚠️  ALGUNS TESTES FALHARAM")
        print("  📋 Verifique as mensagens acima para detalhes")
        print("\n  Testes que falharam:")
        for name, passed in results.items():
            if not passed:
                print(f"     - {name}")

def main():
    """Executa todos os testes"""
    print("\n" + "🔍 SISORC ULTIMATE - DIAGNÓSTICO DO SISTEMA".center(60))
    print("v13.0 - Professional Edition\n")
    
    results = {}
    
    # Executa testes
    results['Importações'] = test_imports()
    results['Estrutura'] = test_structure()
    results['Configuração'] = test_config()
    results['Logger'] = test_logger()
    results['Sanitizer'] = test_sanitizer()
    results['Engine'] = test_engine()
    results['Database'] = test_database()
    results['Interface'] = test_ui()
    
    # Info adicional
    create_example_files()
    
    # Resumo
    print_summary(results)
    
    print("\n" + "="*60)
    
    # Pausa no final
    input("\nPressione ENTER para sair...")

if __name__ == "__main__":
    main()
