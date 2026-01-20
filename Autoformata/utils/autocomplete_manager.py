import json
import os
from pathlib import Path

class AutocompleteManager:
    def __init__(self):
        # Garante caminho absoluto
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.path = Path(base_dir) / "config" / "autocomplete.json"
        
        if not self.path.parent.exists():
            self.path.parent.mkdir(parents=True, exist_ok=True)
            
        self.data = {}
        self._load()

    def _load(self):
        if not self.path.exists():
            self.data = {
                "campus": [], "setor": [], "servidor": [],
                "elaborador": [], "estagiario": [], "fiscal": []
            }
            self.save()
        else:
            try:
                with open(self.path, 'r', encoding='utf-8') as f:
                    raw_data = json.load(f)
                    self.data = {k.lower(): v for k, v in raw_data.items()}
            except (json.JSONDecodeError, Exception) as e:
                print(f"Erro no autocomplete: {e}")
                self.data = {}

    def save(self):
        try:
            with open(self.path, 'w', encoding='utf-8') as f:
                json.dump(self.data, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"Erro ao salvar: {e}")

    def get_list(self, key):
        key = key.lower()
        lista = self.data.get(key, [])
        if isinstance(lista, list):
            # Filtra, converte pra string, remove duplicatas e ordena
            return sorted(list(set([str(x).strip().upper() for x in lista if x and str(x).strip()])))
        return []

    def add_value(self, key, value):
        if not value or not str(value).strip(): return
        
        key = key.lower()
        val_str = str(value).strip().upper()
        
        if key not in self.data: self.data[key] = []
        if not isinstance(self.data[key], list): self.data[key] = []

        if val_str not in self.data[key]:
            self.data[key].append(val_str)
            self.save()

    def remove_value(self, key, value):
        """Remove um item da lista e salva"""
        key = key.lower()
        val_str = str(value).strip().upper()
        
        if key in self.data and val_str in self.data[key]:
            self.data[key].remove(val_str)
            self.save()
            return True
        return False