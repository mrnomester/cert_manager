# config.py
import sys
import os
from pathlib import Path
from datetime import timedelta

class Config:
    # Основные пути
    NETWORK_FOLDER = Path(r"\\nas\DOC\ИТ\сертификаты")
    ARCHIVE_FOLDER = Path(r"\\nas\DOC\ИТ\сертификаты\1 архив")
    EXCEL_FILE = Path(r"X:\ИТ\СОТРУДНИКИ (полный).xlsx")
    LOG_FOLDER = Path(r"\\nas\Distrib\script\certificate\cert_log")

    # Настройки приложения
    CERT_EXPIRY_MONTHS = 15
    CERT_EXPIRY_DAYS = CERT_EXPIRY_MONTHS * 30  # 15 месяцев в днях
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
CRYPTO_PC = "NUC058"
CRYPTO_USER = "crypto_pro"
CRYPTO_PRO_PATH = Path(f"\\\\{CRYPTO_PC}\\c$\\users\\{CRYPTO_USER}\\AppData\\Local\\Crypto Pro")

# Стилевые настройки
DARK_THEME = True
FONT_FAMILY = "Segoe UI"
FONT_SIZE = 12

# Настройки LAPS
LAPS_MODULE_PATH = "AdmPwd.PS"

config = Config()