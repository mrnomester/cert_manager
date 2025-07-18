# ui/search_view.py
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton,
    QListWidget, QGroupBox, QMessageBox
)
from PySide6.QtGui import QIcon, QColor
from PySide6.QtCore import Qt
from pathlib import Path
from core.certificates import CertificateManager
from config import CRYPTO_PRO_PATH
import os
from config import resource_path

class SearchView(QWidget):
    def __init__(self):
        super().__init__()
        self.cert_manager = CertificateManager()
        self.setup_ui()
        self.setup_connections()
        
    def setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(8)

        # Стили элементов
        element_style = """
            QPushButton {
                min-height: 28px;
                padding: 4px 8px;
                background-color: #333;
                border: 1px solid #444;
                border-radius: 4px;
                color: #eee;
            }
            QPushButton:hover {
                background-color: #3a3a3a;
            }
            QLineEdit {
                background-color: #2d2d2d;
                border: 1px solid #3a3a3a;
                padding: 5px 8px;
                border-radius: 3px;
                color: #fff;
                min-height: 28px;
            }
            QGroupBox {
                margin-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 5px;
                padding: 0 3px;
                color: #e0e0e0;
            }
        """

        # Блок поиска
        search_group = QGroupBox("Поиск")
        search_group.setStyleSheet(element_style)
        search_layout = QVBoxLayout()
        search_layout.setContentsMargins(8, 15, 8, 8)
        
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Введите клиента/контейнер")
        search_layout.addWidget(self.search_input)
        
        btn_layout = QHBoxLayout()
        self.search_btn = QPushButton(QIcon(resource_path("resources/search.png")), "Найти")
        self.client_btn = QPushButton(QIcon(resource_path("resources/client.png")), "Клиенты")
        self.cert_btn = QPushButton(QIcon(resource_path("resources/cert.png")), "Сертификаты")
        
        for btn in [self.search_btn, self.client_btn, self.cert_btn]:
            btn.setStyleSheet(element_style)
            btn_layout.addWidget(btn)
        
        search_layout.addLayout(btn_layout)
        search_group.setLayout(search_layout)
        layout.addWidget(search_group)

        # Блок результатов
        result_group = QGroupBox("Результаты")
        result_group.setStyleSheet(element_style)
        result_layout = QVBoxLayout()
        result_layout.setContentsMargins(8, 15, 8, 8)
        
        self.result_list = QListWidget()
        self.result_list.setStyleSheet("""
            QListWidget {
                background-color: #252525;
                color: #e0e0e0;
                border: 1px solid #3a3a3a;
                font-size: 13px;
            }
            QListWidget::item {
                padding: 5px;
            }
            QListWidget::item:selected {
                background-color: #3a3a3a;
            }
        """)
        result_layout.addWidget(self.result_list)
        
        self.delete_btn = QPushButton(QIcon(resource_path("resources/delete.png")), "Удалить")
        self.delete_btn.setStyleSheet("""
            QPushButton {
                background-color: #7a2a2a;
                color: white;
            }
            QPushButton:hover {
                background-color: #8a3232;
            }
        """)
        result_layout.addWidget(self.delete_btn)
        
        result_group.setLayout(result_layout)
        layout.addWidget(result_group)

        self.setLayout(layout)

    def setup_connections(self):
        """Настройка сигналов и слотов"""
        self.search_btn.clicked.connect(self.search_items)
        self.client_btn.clicked.connect(self.search_clients)
        self.cert_btn.clicked.connect(self.search_certs)
        self.delete_btn.clicked.connect(self.delete_selected)
        
    def search_items(self):
        """Общий поиск по клиентам и сертификатам"""
        query = self.search_input.text().strip()
        if not query:
            self.show_message("Ошибка", "Введите запрос для поиска")
            return
            
        clients = self.cert_manager.find_clients(query)
        certs = self.search_certificates(query)
        
        if not clients and not certs:
            self.show_message("Результат", "Ничего не найдено")
            return
            
        self.result_list.clear()
        
        if clients:
            self.result_list.addItem("=== Клиенты ===")
            for client in clients:
                self.result_list.addItem(f"Клиент: {client.name}")
        
        if certs:
            self.result_list.addItem("\n=== Сертификаты ===")
            for cert in certs:
                status = "✓" if cert["status"] == "valid" else "✗"
                self.result_list.addItem(f"{cert['name']} {status} ({cert['date']})")

    def search_clients(self):
        """Поиск только клиентов"""
        query = self.search_input.text().strip()
        if not query:
            self.show_message("Ошибка", "Введите имя клиента")
            return
            
        clients = self.cert_manager.find_clients(query)
        if not clients:
            self.show_message("Результат", "Клиенты не найдены")
            return
            
        self.result_list.clear()
        self.result_list.addItem("Найденные клиенты:")
        for client in clients:
            self.result_list.addItem(f"• {client.name}")

    def search_certs(self):
        """Поиск сертификатов"""
        query = self.search_input.text().strip()
        if not query:
            self.show_message("Ошибка", "Введите название сертификата")
            return
            
        certs = self.search_certificates(query)
        if not certs:
            self.show_message("Результат", "Сертификаты не найдены")
            return
            
        self.result_list.clear()
        self.result_list.addItem("Найденные сертификаты:")
        for cert in certs:
            status = "✓" if cert["status"] == "valid" else "✗"
            self.result_list.addItem(f"• {cert['name']} {status} ({cert['date']})")

    def search_certificates(self, query: str) -> list:
        """Поиск сертификатов по запросу"""
        clients = self.cert_manager.find_clients("")
        certs = []
        for client in clients:
            for cert in self.cert_manager.get_certificates(client):
                if query.lower() in cert["name"].lower():
                    certs.append({
                        "path": cert["path"],
                        "name": f"{client.name}/{cert['name']}",
                        "status": cert["status"],
                        "date": cert["date"]
                    })
        return certs

    def delete_selected(self):
        """Удаление выбранного сертификата"""
        selected_item = self.result_list.currentItem()
        if not selected_item or not selected_item.text().startswith(("•", "Клиент:")):
            self.show_message("Ошибка", "Выберите сертификат для удаления")
            return
            
        text = selected_item.text()
        if text.startswith("Клиент:"):
            self.show_message("Информация", "Удаление клиентов не поддерживается")
            return
            
        # Парсим имя сертификата
        cert_name = text.split()[1].split("/")[-1]
        cert_path = Path(CRYPTO_PRO_PATH) / cert_name
        
        # Подтверждение удаления
        reply = QMessageBox.question(
            self, "Подтверждение",
            f"Удалить сертификат {cert_name}?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            success, message = self.cert_manager.delete_certificate(cert_path)
            if success:
                self.result_list.takeItem(self.result_list.row(selected_item))
                self.show_message("Успех", message)
            else:
                self.show_message("Ошибка", message)
                
    def show_message(self, title, message):
        """Отображение информационного сообщения"""
        msg = QMessageBox(self)
        msg.setWindowTitle(title)
        msg.setText(message)
        
        # Рассчитываем оптимальную ширину
        font_metrics = msg.fontMetrics()
        max_width = max(font_metrics.horizontalAdvance(line) for line in message.split('\n'))
        msg.setMinimumWidth(min(max_width + 50, 800))  # Ограничиваем максимальную ширину
        
        msg.setStyleSheet("""
            QMessageBox {
                background-color: #252525;
            }
            QLabel {
                color: #e0e0e0;
            }
        """)
        msg.exec_()