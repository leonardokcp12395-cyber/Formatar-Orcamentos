import json
import os
from pathlib import Path

class ConfigManager:
    def __init__(self):
        self.path = Path("config/profiles.json")
        if not self.path.parent.exists():
            self.path.parent.mkdir(parents=True, exist_ok=True)
        self._check_file()

    def _check_file(self):
        if not self.path.exists():
            self.save_profiles({
                "ultimo_perfil": "PADRAO",
                "perfis": {
                    "PADRAO": {
                        "input": {"ITEM": "ITEM", "DESCRICAO": "DESCRIÇÃO", "UNID": "UND", "QUANT": "QUANT.", "UNIT": "VALOR UNIT"},
                        "output": {"ITEM": "A", "DESCRICAO": "D", "UNID": "E", "QUANT": "F", "UNIT": "G", "TOTAL": "H"}
                    }
                }
            })

    def load_profiles(self):
        try:
            with open(self.path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {"perfis": {}}

    def save_profiles(self, data):
        with open(self.path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)