# Planify v4.0 - Documentação Técnica

## 📋 Visão Geral

Sistema profissional de automação de orçamentos de engenharia civil, construído com arquitetura modular MVC, seguindo princípios SOLID e Clean Architecture.

## 🏗️ Arquitetura do Projeto

```
Planify/
│
├── main.py                      # Entry point (verifica deps e inicia app)
│
├── config/
│   ├── settings.json           # Configurações completas do sistema
│   ├── profiles.json           # Perfis de mapeamento
│   ├── autocomplete.json       # Dados de autocompletar
│   └── templates/              # Modelos Excel importados
│
├── core/                       # Módulos principais (lógica de negócio)
│   ├── sanitizer.py           # Limpeza e blindagem de arquivos Excel
│   ├── excel_handler.py       # Engine de processamento de orçamentos
│   ├── database.py            # Gerenciamento SQLite (histórico)
│   └── paths.py               # Resolução de caminhos (dev/exe)
│
├── controllers/               # Controladores MVC
│   └── main_controller.py     # Controller principal (thread-safe via Queue)
│
├── ui/                        # Interface gráfica
│   ├── main_window.py         # Janela principal (CustomTkinter) — Orquestrador
│   └── components/            # Componentes encapsulados
│       ├── side_panel.py      # Painel lateral (Dados da Obra)
│       ├── config_panel.py    # Aba de Configurações & Mapeamento
│       ├── top_dashboard.py   # Barra superior (Seleção + Importação)
│       └── excel_preview.py   # Tabela de classificação de níveis
│
├── utils/                     # Utilitários
│   ├── logger.py             # Sistema de logging customizado
│   ├── helpers.py            # Funções auxiliares e validadores
│   ├── smart_parser.py       # Extração inteligente de texto (WhatsApp)
│   ├── autocomplete_manager.py # Gestão de listas de autocompletar
│   ├── config_manager.py     # Gestão de perfis de configuração  
│   ├── template_manager.py   # Gestão de modelos Excel
│   └── pdf_exporter.py       # Conversão Excel → PDF via Win32COM
│
├── assets/                    # Recursos (ícones, imagens)
│   └── icon.ico              # Ícone da aplicação
│
├── build_rapido.py            # Script de compilação (PyInstaller)
│
└── dist/                     # Executável compilado (gerado)
    └── Planify.exe
```

## 🔧 Instalação e Configuração

### 1. Dependências

```bash
pip install -r requirements.txt
```

### 2. Primeira Execução

```bash
python main.py
```

### 3. Compilação para Executável

```bash
python build_rapido.py
```

O executável estará em `dist/Planify/Planify.exe`

## 📦 Arquitetura MVC (v4.0)

### Princípios Arquitectónicos

1. **Encapsulamento de Componentes UI**: Cada componente (SidePanel, ConfigPanel, TopDashboard, LevelSelector) é dono dos seus próprios widgets. Nunca injeta referências no parent.
2. **APIs Limpas**: Componentes expõem `get_data()`, `set_data()`, `limpar_campos()` — o orquestrador (main_window) só usa estas APIs.
3. **Thread-Safety via Queue**: O Controller não tem referência à View. Toda comunicação background→UI passa por `queue.Queue`, com callbacks executados na Main Thread.
4. **Garbage Collection Explícita**: Após leitura do Pandas e operações Win32COM, os DataFrames são destruídos e `gc.collect()` é chamado.
5. **Win32COM Blindado**: Todas as chamadas Excel garantem `excel.Quit()` + `del` no bloco `finally`.

### Fluxo de Dados

```
┌──────────────────────────────────────────────────┐
│  PlanifyApp (main_window.py) — ORQUESTRADOR      │
│                                                    │
│  ┌─────────────┐ ┌──────────────┐ ┌────────────┐ │
│  │ SidePanel    │ │ ConfigPanel  │ │ TopDashboard│ │
│  │ get_data()   │ │ get_data()   │ │ get_text()  │ │
│  │ set_data()   │ │ get_mapping()│ │ set_label() │ │
│  │ limpar()     │ │ update_cols()│ │ clear()     │ │
│  └──────┬───────┘ └──────┬───────┘ └──────┬──────┘ │
│         │                │                │         │
│         └────────────────┼────────────────┘         │
│                          │                          │
│    ┌─────────────────────▼──────────────────────┐   │
│    │       MainController (queue.Queue)          │   │
│    │  ✗ NÃO tem self.view                       │   │
│    │  ✓ Publica eventos na queue com _handler   │   │
│    └─────────────────────┬──────────────────────┘   │
│                          │                          │
└──────────────────────────┼──────────────────────────┘
                           │
          ┌────────────────▼────────────────┐
          │  OrcamentoEngine + Win32COM     │
          │  (Background Threads)           │
          │  progress_callback → Queue      │
          └─────────────────────────────────┘
```

## 📦 Módulos Principais

### `core/excel_handler.py` - Core Engine

**Classe: `OrcamentoEngine`**

- Leitura inteligente de dados com Pandas
- Classificação automática: Título vs. Item
- Mapeamento dinâmico de colunas
- Progress Tracker Real (0-100%) via callback
- Conversor `_parse_num` seguro (deteta float/int nativo do Pandas)

### `controllers/main_controller.py` - Controller MVC

- **Sem referência directa à UI** — recebe `queue.Queue` e `schedule_fn`
- Background threads publicam eventos na queue com `_handler` callback
- Garbage Collection após operações Pandas
- Win32COM com cleanup garantido no `finally`

### `ui/components/` - Componentes Encapsulados

| Componente | Responsabilidade | API Principal |
|---|---|---|
| `SidePanel` | Dados da Obra | `get_data()`, `set_data()`, `limpar_campos()` |
| `ConfigPanel` | Configurações & Mapeamento | `get_data()`, `get_column_mapping()`, `update_column_options()` |
| `TopDashboard` | Seleção ficheiro + WhatsApp | `get_import_text()`, `set_file_label()` |
| `LevelSelector` | Tabela de classificação | `add_row()`, `get_final_data()`, `clear()` |

### `utils/smart_parser.py` - Parser Inteligente

- Usa `rapidfuzz` para fuzzy matching
- Regex do `valor_simulado`: `r"(?:VALOR|ORÇAMENTO|ESTIMATIVA|TOTAL)\s*[:]?\s*(?:R\$)?\s*([\d\.,]+)"`
- Normalização automática contra histórico do autocomplete

## 🔍 Fluxo de Processamento

```
1. [UI] Usuário seleciona/arrasta arquivos e preenche dados
         ↓
2. [Win32COM] Limpa arquivo Excel problemático (background thread)
         ↓ (queue → Main Thread)
3. [Pandas] Lê dados do arquivo limpo
         ↓ (queue → Main Thread)
4. [UI] Usuário classifica linhas (N1, N2, N3, ITEM, IGNORAR)
         ↓
5. [Engine] Processa itens com progress real (0-100%)
         ↓ (queue → Main Thread)
6. [OpenPyXL] Escreve arquivo final com fórmulas
         ↓
7. [Database] Salva registro no histórico
         ↓
8. [UI] Exibe sucesso e abre o arquivo
```

## ⚙️ Correções Técnicas v4.0

### 1. ttk.Style — Nome do estilo
**Bug**: Nome `"Treeview"` ou `"ExcelTreeview"` não herda layout base do Tkinter.
**Fix**: Usa `"Excel.Treeview"` (contém `.Treeview`).

### 2. Âncora de colunas
**Bug**: `anchor="right"` — Tkinter não suporta.
**Fix**: `anchor="e"` (East).

### 3. `_parse_num` — Conversão segura
**Bug**: Sempre convertia para string antes de calcular, mesmo quando o Pandas já dava `float`.
**Fix**: Verifica `isinstance(val, (int, float))` e preserva.

### 4. Thread-Safety
**Bug**: Controller acedia `self.view._on_...()` directamente.
**Fix**: Controller publica na queue com `_handler` callback, View processa na Main Thread.

### 5. Win32COM — Processos Fantasma
**Bug**: `excel.Quit()` nem sempre era chamado em caso de excepção.
**Fix**: Bloco `finally` com `del wb`, `del excel`, `gc.collect()`.

## 🐛 Troubleshooting

### ❌ "Não foi possível localizar o cabeçalho"
Abra o sintético no Excel e verifique se há uma linha com "ITEM" e "DESCRIÇÃO".

### ❌ "Arquivo de saída está aberto"
Feche todos os arquivos Excel gerados pelo Planify.

### ❌ "Nenhum dado encontrado"
Ajuste "Linha Inicial" na aba Configurações e recarregue.

## 📄 Licença e Créditos

**Planify v4.0**
Desenvolvido para automatizar orçamentos de obras públicas.

---

✅ **Sistema completamente refatorado com arquitetura MVC limpa**
🚀 **Pronto para uso em produção**
📊 **Thread-safe e memory-optimized**