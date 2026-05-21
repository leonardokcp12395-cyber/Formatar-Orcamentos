import json
import os
from pathlib import Path
from pydantic import BaseModel, Field
from typing import Dict, Optional

class ProfileMapping(BaseModel):
    input: Dict[str, str] = Field(default_factory=dict)
    output: Dict[str, str] = Field(default_factory=dict)

class ConfigSchema(BaseModel):
    ultimo_perfil: str = "PADRAO"
    perfis: Dict[str, ProfileMapping] = Field(
        default_factory=lambda: {
            "PADRAO": ProfileMapping(
                input={"ITEM": "ITEM", "DESCRICAO": "DESCRIÇÃO", "UNID": "UND", "QUANT": "QUANT.", "UNIT": "VALOR UNIT"},
                output={"ITEM": "A", "DESCRICAO": "D", "UNID": "E", "QUANT": "F", "UNIT": "G", "TOTAL": "H"}
            )
        }
    )

class ConfigManager:
    def __init__(self):
        self.path = Path("config/profiles.json")
        if not self.path.parent.exists():
            self.path.parent.mkdir(parents=True, exist_ok=True)
        self._check_file()

    def _check_file(self):
        if not self.path.exists():
            default_config = ConfigSchema()
            self.save_profiles(default_config.model_dump())

    def load_profiles(self) -> dict:
        try:
            with open(self.path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            # Valida silenciosamente a estrutura. Se quebrar, fallback.
            validated = ConfigSchema(**data)
            return validated.model_dump()
        except:
            from utils.logger import Logger
            Logger.warning("Config corrompida! Retornando defaults salvos.")
            default_config = ConfigSchema()
            # Salva o arquivo default novamente
            self.save_profiles(default_config.model_dump())
            return default_config.model_dump()

    def save_profiles(self, data: dict):
        try:
            validated = ConfigSchema(**data)
            with open(self.path, 'w', encoding='utf-8') as f:
                json.dump(validated.model_dump(), f, indent=4, ensure_ascii=False)
        except Exception as e:
            from utils.logger import Logger
            Logger.error(f"Erro ao salvar configurações do perfil: {e}")