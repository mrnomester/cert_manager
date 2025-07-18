# core/employees.py
import pandas as pd
from pathlib import Path
from typing import List, Dict, Optional
import time
from functools import lru_cache
from config import config

class EmployeeManager:
    def __init__(self, excel_path: Path = config.EXCEL_FILE):
        self.excel_path = excel_path
        self._cache = None
        self._last_load_time = 0
        
    def _load_dataframe(self) -> pd.DataFrame:
        """Загрузка данных из Excel в DataFrame"""
        try:
            df = pd.read_excel(self.excel_path, dtype=str)
            df.fillna("", inplace=True)
            return df
        except Exception as e:
            raise ValueError(f"Ошибка загрузки Excel: {e}")

    def _validate_dataframe(self, df: pd.DataFrame) -> bool:
        """Проверка структуры DataFrame"""
        required_fields = ['Фамилия', 'ИО', 'ПК', 'Username']
        missing = [field for field in required_fields if field not in df.columns]
        if missing:
            raise ValueError(f"Отсутствуют обязательные поля: {', '.join(missing)}")
        return True

    def _load_employees(self) -> List[Dict]:
        """Загрузка и валидация данных"""
        df = self._load_dataframe()
        self._validate_dataframe(df)
        self._cache = df.to_dict('records')
        self._last_load_time = time.time()
        return self._cache

    @property
    def employees(self) -> List[Dict]:
        """Получение актуального списка сотрудников"""
        return self._load_employees()

    @staticmethod
    def normalize_name(name: str) -> str:
        """Нормализация имени для поиска"""
        return ' '.join(name.strip().split()).lower()

    def search_by_field(self, query: str, field: str = 'Фамилия') -> List[Dict]:
        """Поиск по конкретному полю"""
        query = self.normalize_name(query)
        return [emp for emp in self.employees 
                if query in self.normalize_name(emp.get(field, ""))]

    def search(self, query: str) -> List[Dict]:
        """Основной метод поиска сотрудников"""
        query = self.normalize_name(query)
        results = []
        
        for emp in self.employees:
            last_name = self.normalize_name(emp.get("Фамилия", ""))
            full_name = self.normalize_name(f"{emp.get('Фамилия', '')} {emp.get('ИО', '')}")
            
            if (query in last_name) or (query == full_name):
                results.append(emp)
                
        return results

    @staticmethod
    def format_employee_info(employee: Dict) -> str:
        """Форматирование информации о сотруднике"""
        return (
            f"{employee.get('Фамилия', '')} {employee.get('ИО', '')} | "
            f"ВН: {employee.get('ВН', 'N/A')} | "
            f"Каб: {employee.get('Каб.', 'N/A')} | "
            f"ПК: {employee.get('ПК', 'N/A')} | "
            f"User: {employee.get('Username', 'N/A')}"
        )

    @staticmethod
    def get_crypto_path(pc_name: str, username: str) -> Path:
        """Генерация пути к папке Crypto Pro"""
        if not pc_name or not username:
            raise ValueError("Не указано имя ПК или пользователя")
        return Path(f"\\\\{pc_name}\\c$\\users\\{username}\\AppData\\Local\\Crypto Pro")