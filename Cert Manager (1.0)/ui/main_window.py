# ui/main_window.py
from PySide6.QtWidgets import QMainWindow, QTabWidget, QStatusBar
from PySide6.QtGui import QIcon
from ui.copy_view import CopyView
from ui.search_view import SearchView
from config import resource_path

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CertManager Pro")
        self.setFixedSize(1000, 700)  # Фиксированный размер
        self.setWindowIcon(QIcon(resource_path("resources/app_icon.png")))
        self.setup_ui()
        self.apply_styles()

    def setup_ui(self):
        # Создаем виджеты вкладок
        self.tabs = QTabWidget()
        self.copy_tab = CopyView()
        self.search_tab = SearchView()
        
        # Добавляем вкладки
        self.tabs.addTab(self.copy_tab, "Копирование")
        self.tabs.addTab(self.search_tab, "Поиск")
        
        # Настройка центрального виджета
        self.setCentralWidget(self.tabs)
        
        # Статус бар
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Готово")

    def apply_styles(self):
        from config import resource_path
        self.setStyleSheet("""
            /* Основные стили */
            QMainWindow {
                background-color: #1e1e1e;
            }
            
            /* Общие стили для всех виджетов */
            QWidget {
                color: #e0e0e0;
                font-family: 'Segoe UI';
                font-size: 14px;
                selection-background-color: #3a3a3a;
            }
            
            /* Вкладки */
            QTabWidget::pane {
                border: 1px solid #3a3a3a;
                background: #252525;
            }
            
            QTabBar::tab {
                background: #252525;
                border: 1px solid #3a3a3a;
                border-bottom: none;
                padding: 8px 15px;
                min-width: 100px;
            }
            
            QTabBar::tab:selected {
                background: #353535;
                border-color: #3a3a3a;
            }
            
            QTabBar::tab:hover {
                background: #303030;
            }
            
            /* Группы */
            QGroupBox {
                border: 1px solid #3a3a3a;
                background-color: #252525;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 20px;
            }
            
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
            
            /* Кнопки - основная стилизация */
            QPushButton {
                background-color: #333333;
                border: 1px solid #444444;
                color: #e0e0e0;
                padding: 5px 10px;
                min-height: 24px;
                border-radius: 3px;
                min-width: 70px;
            }
            
            QPushButton:hover {
                background-color: #3a3a3a;
                border: 1px solid #555555;
            }
            
            QPushButton:pressed {
                background-color: #2a2a2a;
            }
            
            QPushButton:disabled {
                background-color: #282828;
                color: #888;
            }
            
            /* Специальные кнопки */
            QPushButton#execute_btn {
                background-color: #2a5699;
                font-weight: bold;
            }
            
            QPushButton#execute_btn:hover {
                background-color: #2e63b0;
            }
            
            QPushButton#delete_btn {
                background-color: #7a2a2a;
            }
            
            QPushButton#delete_btn:hover {
                background-color: #8a3232;
            }
            
            /* Поля ввода */
            QLineEdit {
                background-color: #2d2d2d;
                border: 1px solid #3a3a3a;
                padding: 5px 8px;
                border-radius: 3px;
            }
            
            QLineEdit:focus {
                border: 1px solid #555555;
            }
            
            /* Чекбоксы */
            QCheckBox {
                spacing: 5px;
            }
            
            QCheckBox::indicator {
                width: 16px;
                height: 16px;
            }
            
            /* Статус бар */
            QStatusBar {
                background-color: #252525;
                border-top: 1px solid #3a3a3a;
                padding: 3px;
            }
            
            QStatusBar::item {
                border: none;
            }
            
            /* Выпадающие списки */
            QComboBox {
                background-color: #2d2d2d;
                border: 1px solid #3a3a3a;
                padding: 3px;
                min-width: 100px;
            }
            
            QComboBox::drop-down {
                width: 20px;
                border-left: 1px solid #3a3a3a;
            }
            
            /* Стили полосы прокрутки */
            QScrollBar:vertical {
                border: none;
                background: #252525;
                width: 10px;
                margin: 0;
            }
            
            QScrollBar::handle:vertical {
                background: #aaaaaa;
                min-height: 20px;
                border-radius: 5px;
            }
            
            QScrollBar::add-line:vertical, 
            QScrollBar::sub-line:vertical {
                height: 0;
                background: none;
            }
            
            QScrollBar::add-page:vertical, 
            QScrollBar::sub-page:vertical {
                background: none;
            }
            
            /* Стили для QMessageBox и QDialog */
            QMessageBox {
                background-color: #252525;
                color: #e0e0e0;
            }
            QMessageBox QLabel {
                color: #e0e0e0;
            }
            QMessageBox QPushButton {
                background-color: #3a3a3a;
                color: #e0e0e0;
                min-width: 80px;
                padding: 5px;
            }
            QDialog {
                background-color: #252525;
                color: #e0e0e0;
            }
        """)

    def closeEvent(self, event):
        """Обработчик закрытия окна"""
        # Можно добавить сохранение настроек или проверку перед закрытием
        event.accept()