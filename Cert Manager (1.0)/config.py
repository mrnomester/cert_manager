# config.py
import sys
import os
from pathlib import Path
from datetime import timedelta

class Config:
    # Основные пути
    NETWORK_FOLDER = Path(r" ") # Хранилище клентов
    ARCHIVE_FOLDER = Path(r" ") # Архив клиентов
    EXCEL_FILE = Path(r" ") # Excel база сотрудников
    LOG_FOLDER = Path(r" ") # Логи установки сертификатов

    # Настройки приложения
    CERT_EXPIRY_DAYS = 490  # ~15 месяцев в днях
    MAX_LOG_ENTRIES = 1000 

    @classmethod
    def validate_paths(cls):
        """Проверка доступности всех путей"""
        results = {}
        for name, path in vars(cls).items():
            if isinstance(path, Path):
                try:
                    results[name] = path.exists()
                except Exception:
                    results[name] = False
        return results



# Настройки Crypto Pro 
CRYPTO_PC = "NUC058" # ПК ответсвенного за сертификаты чтобы удалять их
CRYPTO_USER = "crypto_pro" # ответсвенный за сертификаты пользователь
CRYPTO_PRO_PATH = Path(f"\\\\{CRYPTO_PC}\\c$\\users\\{CRYPTO_USER}\\AppData\\Local\\Crypto Pro")

# Стилевые настройки
DARK_THEME = True
FONT_FAMILY = "Segoe UI"
FONT_SIZE = 12

# Настройки LAPS
LAPS_MODULE_PATH = "AdmPwd.PS"

config = Config()