import json
import os
import shutil
from pathlib import Path
from core.paths import get_app_dir

class TemplateManager:
    def __init__(self):
        # Pasta onde os modelos físicos ficarão guardados
        self.models_dir = get_app_dir() / "config" / "templates"
        self.config_file = self.models_dir / "templates.json"
        
        self._ensure_structure()
        self.templates = self._load_templates()

    def _ensure_structure(self):
        if not self.models_dir.exists():
            self.models_dir.mkdir(parents=True, exist_ok=True)
        
        # Se não tiver arquivo de config, cria um padrão vazio
        if not self.config_file.exists():
            self._save_config({})

    def _load_templates(self):
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}

    def _save_config(self, data):
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)

    def get_template_names(self):
        """Retorna lista de nomes para o ComboBox"""
        return sorted(list(self.templates.keys()))

    def get_template_path(self, name):
        """Retorna o caminho absoluto do arquivo Excel do modelo"""
        if name in self.templates:
            filename = self.templates[name]['filename']
            return str(self.models_dir / filename)
        return None

    def get_template_info(self, name):
        return self.templates.get(name, {})

    def add_template(self, name, source_path, start_line=25):
        """Importa um novo modelo para o sistema"""
        if not os.path.exists(source_path):
            return False, "Arquivo de origem não encontrado."

        # Copia o arquivo para a pasta segura do sistema
        filename = f"{name.replace(' ', '_')}.xlsx"
        dest_path = self.models_dir / filename
        
        try:
            shutil.copy2(source_path, dest_path)
            
            # Salva no JSON
            self.templates[name] = {
                "filename": filename,
                "start_line": int(start_line),
                "date_added": str(os.path.getmtime(dest_path))
            }
            self._save_config(self.templates)
            return True, "Modelo importado com sucesso!"
        except Exception as e:
            return False, f"Erro ao importar: {e}"

    def remove_template(self, name):
        """Remove o modelo do JSON e deleta o arquivo"""
        if name in self.templates:
            filename = self.templates[name]['filename']
            file_path = self.models_dir / filename
            
            # Remove do JSON
            del self.templates[name]
            self._save_config(self.templates)
            
            # Tenta remover o arquivo físico
            try:
                if file_path.exists():
                    os.remove(file_path)
            except: pass 
            
            return True
        return False