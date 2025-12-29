"""
SISORC ULTIMATE - Build Script
Script para compilar aplicação em executável usando PyInstaller
"""

import subprocess
import sys
import os
from pathlib import Path

def verificar_pyinstaller():
    """Verifica se PyInstaller está instalado"""
    try:
        import PyInstaller
        print("✅ PyInstaller encontrado")
        return True
    except ImportError:
        print("❌ PyInstaller não encontrado")
        print("📦 Instalando PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        return True

def criar_spec_file():
    """Cria arquivo .spec customizado para PyInstaller"""
    
    spec_content = """
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('config/settings.json', 'config'),
    ],
    hiddenimports=[
        'customtkinter',
        'pandas',
        'openpyxl',
        'sqlite3',
        'PIL._tkinter_finder'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='SISORC_ULTIMATE',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # False = Sem console (GUI pura)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/icon.ico'  # Adicione um ícone aqui
)
"""
    
    with open('sisorc.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content)
    
    print("✅ Arquivo sisorc.spec criado")

def compilar():
    """Compila aplicação usando PyInstaller"""
    print("\n🔨 Iniciando compilação...")
    print("⏳ Este processo pode levar alguns minutos...\n")
    
    # Comando PyInstaller
    comando = [
        "pyinstaller",
        "--clean",  # Limpa build anterior
        "--noconfirm",  # Não pede confirmação
        "sisorc.spec"
    ]
    
    try:
        resultado = subprocess.run(comando, check=True)
        return resultado.returncode == 0
    except subprocess.CalledProcessError as e:
        print(f"\n❌ Erro na compilação: {e}")
        return False

def criar_estrutura_dist():
    """Cria estrutura de pastas no dist"""
    dist_path = Path("dist")
    
    if not dist_path.exists():
        print("❌ Pasta dist não encontrada")
        return False
    
    # Cria pasta config se não existir
    config_dist = dist_path / "config"
    config_dist.mkdir(exist_ok=True)
    
    # Copia settings.json se ainda não estiver lá
    settings_origem = Path("config/settings.json")
    settings_destino = config_dist / "settings.json"
    
    if settings_origem.exists() and not settings_destino.exists():
        import shutil
        shutil.copy(settings_origem, settings_destino)
        print("✅ Arquivo settings.json copiado para dist")
    
    return True

def criar_readme():
    """Cria README de distribuição"""
    readme_content = """
# SISORC ULTIMATE - Sistema de Orçamentação Automatizado

## 🚀 Guia de Uso Rápido

### Primeira Execução
1. Execute o arquivo `SISORC_ULTIMATE.exe`
2. Aguarde a inicialização (pode levar alguns segundos)

### Gerando um Orçamento
1. Vá para a aba "🏗️ Gerador"
2. Clique em "📊 Selecionar Planilha Sintética" e escolha seu arquivo de dados
3. Clique em "📄 Selecionar Modelo" e escolha o template Excel
4. Preencha os dados do projeto (Obra, Local, BDI)
5. Clique em "🚀 EXECUTAR AUTOMAÇÃO"
6. Aguarde o processamento
7. O arquivo final será salvo na mesma pasta do executável

### Recursos
- **Histórico**: Veja todos os orçamentos gerados
- **Console**: Acompanhe logs detalhados em tempo real
- **Configurações**: Personalize o tema e outras opções

### Requisitos do Arquivo de Entrada
- Formato: Excel (.xlsx ou .xls)
- Deve conter as colunas:
  - ITEM
  - DESCRIÇÃO DO SERVIÇO
  - UNID.
  - QUANTID.
  - PREÇO UNIT.(SEM BDI)
  - PREÇO UNIT.(COM BDI)

### Suporte
Para problemas ou dúvidas, verifique o arquivo `sisorc_log.txt` gerado automaticamente.

---
Desenvolvido por Engineering Automation Lab
Versão 3.0.0
"""
    
    with open("dist/README.txt", "w", encoding='utf-8') as f:
        f.write(readme_content)
    
    print("✅ README.txt criado")

def main():
    """Função principal de build"""
    print("=" * 60)
    print("  SISORC ULTIMATE - BUILD SCRIPT")
    print("=" * 60)
    print()
    
    # 1. Verifica PyInstaller
    if not verificar_pyinstaller():
        print("❌ Falha ao instalar PyInstaller")
        return False
    
    # 2. Cria arquivo .spec
    criar_spec_file()
    
    # 3. Compila
    sucesso = compilar()
    
    if sucesso:
        print("\n" + "=" * 60)
        print("  ✅ COMPILAÇÃO CONCLUÍDA COM SUCESSO!")
        print("=" * 60)
        print()
        
        # 4. Configura distribuição
        criar_estrutura_dist()
        criar_readme()
        
        print("\n📦 Arquivos prontos para distribuição:")
        print("   📂 dist/SISORC_ULTIMATE.exe")
        print("   📂 dist/config/settings.json")
        print("   📂 dist/README.txt")
        print()
        print("🎉 Você já pode distribuir a pasta 'dist' completa!")
        print()
        
    else:
        print("\n" + "=" * 60)
        print("  ❌ COMPILAÇÃO FALHOU")
        print("=" * 60)
        print()
        print("Verifique os erros acima e tente novamente.")
    
    return sucesso

if __name__ == "__main__":
    sucesso = main()
    
    input("\nPressione ENTER para sair...")
    sys.exit(0 if sucesso else 1)