# ui/base_view.py
from PySide6.QtWidgets import QWidget
from config import config

class BaseView(QWidget):
    """Базовый класс для всех view с общими методами"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.config = config
        self.setup_ui()
        self.setup_connections()
        
    def setup_ui(self):
        """Настройка интерфейса (должен быть переопределен)"""
        raise NotImplementedError
        
    def setup_connections(self):
        """Настройка сигналов и слотов (должен быть переопределен)"""
        raise NotImplementedError
        
    def validate_input(self, value: str, min_length: int = 2) -> bool:
        """Базовая валидация ввода"""
        return len(value.strip()) >= min_length