import psutil
from utils.logger import Logger

def clean_zombie_excels(force=False):
    """
    Lista e força encerramento das instâncias de EXCEL.EXE na máquina local.
    """
    cleaned_count = 0
    if not force:
        return 0 # Só limpamos sob demanda (panic button) para não matar o excel pessoal do user à toa na inicialização.
        
    try:
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                name = proc.info.get('name', '').lower()
                if name == 'excel.exe':
                    proc.kill()
                    cleaned_count += 1
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
                
        if cleaned_count > 0:
            Logger.warning(f"🚨 BOTÃO DE PÂNICO: {cleaned_count} processo(s) EXCEL.EXE foram finalizados à força!")
            
        return cleaned_count
    except Exception as e:
        Logger.error(f"Erro ao limpar Excel: {e}")
        return 0
