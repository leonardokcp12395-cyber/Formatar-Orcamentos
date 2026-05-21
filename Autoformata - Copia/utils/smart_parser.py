import re
from rapidfuzz import process, fuzz


class SmartParser:
    """
    V2: Parser com Normalização de Dados.
    Usa Regex para capturar e Fuzzy Matching (via RapidFuzz) para corrigir baseando-se no histórico.
    """

    @staticmethod
    def parse_whatsapp_text(text: str, autocomplete_mgr=None) -> dict:
        """
        Analisa texto, extrai campos e (se fornecido o gerenciador) padroniza os valores.
        """
        data = {}

        # 1. Limpeza básica pré-processamento
        clean_text = text.replace('*', '').replace(':', ' : ')
        lines = clean_text.split('\n')

        # 2. Dicionário de Padrões
        patterns = {
            "campus": r"(?:CAMPUS|CAMPI|UNIDADE)\s*[:]?\s*(.*)",
            "setor": r"(?:SETOR|LOCAL|DEPARTAMENTO|COORDENAÇÃO)\s*[:]?\s*(.*)",
            "descricao_header": r"(?:DESCRIÇÃO|DISCRIMINAÇÃO|OBJETO|SERVIÇO)\s*[:]?\s*(.*)",
            "servidor": r"(?:SOLICITANTE|DEMANDANTE|REQUISITANTE|SERVIDOR)\s*[:]?\s*(.*)",
            "fiscal": r"(?:FISCAL|FISCALIZAÇÃO)\s*[:]?\s*(.*)",
            "elaborador": r"(?:ORÇAMENTO|ORÇAMENTISTA|RESPONSÁVEL|ELABORADO)\s*[:]?\s*(.*)",
            "estagiario": r"(?:ESTAGIÁRIO|APOIO)\s*[:]?\s*(.*)",
            "processo": r"(?:PROCESSO|SIPAC|SEI|Nº PROCESSO)\s*[:]?\s*([\d\.\-\/]+)",
            "orcafascio": r"(?:ORÇAFASCIO|OF|CÓDIGO)\s*[:]?\s*(\d+)",
            "empenho": r"(?:EMPENHO|NOTA DE EMPENHO)\s*[:]?\s*(.*)",
            "contrato": r"(?:CONTRATO|ATA)\s*[:]?\s*(.*)",
            "num_orcamento": r"(?:ORDEM DE SERVIÇO|OS)(?:\s*(?:N[º°]|NUMERO))?\s*[:\-]?\s*(\d+)",
            "valor_simulado": r"(?:VALOR|ORÇAMENTO|ESTIMATIVA|TOTAL)\s*[:]?\s*(?:R\$)?\s*([\d\.,]+)"
        }

        # 3. Extração via Regex
        for line in lines:
            line = line.strip()
            if not line:
                continue

            for key, pattern in patterns.items():
                match = re.search(pattern, line, re.IGNORECASE)
                if match:
                    val = match.group(1).strip().upper()
                    val = re.sub(r"^(DO |DA |DE |O |A )", "", val)

                    if key == "contrato" and "descricao_header" in data:
                        data["descricao_header"] += f" - {val}"
                    elif key == "contrato":
                        data["_contrato_temp"] = val
                    else:
                        if len(val) > 1:
                            data[key] = val

        # 4. Pós-processamento de lógica
        if "_contrato_temp" in data and "descricao_header" in data:
            data["descricao_header"] += f" - {data['_contrato_temp']}"

        # 5. NORMALIZAÇÃO INTELIGENTE
        if autocomplete_mgr:
            data = SmartParser._normalizar_dados(data, autocomplete_mgr)

        return data

    @staticmethod
    def _normalizar_dados(data, mgr):
        """
        Tenta encontrar o valor mais próximo no banco de dados para corrigir erros de digitação.
        """
        mapa_chaves = {
            "campus": "campus",
            "setor": "setor",
            "servidor": "servidor",
            "elaborador": "elaborador",
            "fiscal": "fiscal"
        }

        for field_parser, field_db in mapa_chaves.items():
            if field_parser in data:
                valor_extraido = data[field_parser]
                opcoes_validas = mgr.get_list(field_db)

                if not opcoes_validas:
                    continue

                if valor_extraido in opcoes_validas:
                    continue

                match = process.extractOne(
                    valor_extraido, opcoes_validas, scorer=fuzz.ratio)

                if match:
                    sugestao, score, _ = match
                    if score >= 60:  # 60 corresponds to cutoff=0.6 from difflib
                        data[field_parser] = sugestao

        return data
