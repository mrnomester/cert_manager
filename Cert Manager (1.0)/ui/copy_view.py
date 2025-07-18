# ui/copy_view.py
import re
import os
import shutil
import subprocess
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from PySide6.QtWidgets import QApplication
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit,
    QCheckBox, QGroupBox, QMessageBox, QCompleter,
    QDialog, QDialogButtonBox, QListWidget, QListWidgetItem
)
from PySide6.QtGui import QIcon, QTextCursor, QColor, QTextCharFormat, QFont, QClipboard
from PySide6.QtCore import Qt, Signal, QStringListModel

from core.employees import EmployeeManager
from core.certificates import CertificateManager
from config import NETWORK_FOLDER, LOG_FOLDER, CERT_EXPIRY_DAYS
from config import resource_path

class CopyView(QWidget):
    log_signal = Signal(str, str)  # Теперь передаём и тип сообщения
    
    def __init__(self):
        super().__init__()
        self.employee_manager = EmployeeManager()
        self.cert_manager = CertificateManager()
        self.setup_ui()
        self.setup_connections()
        self.pending_action = None
        self.log_history = []  # История логов для возможного анализа
        
    def setup_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(5, 5, 5, 5)
        main_layout.setSpacing(5)

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
            }
        """

        # Блок сотрудника
        emp_group = QGroupBox("Сотрудник")
        emp_group.setStyleSheet(element_style)
        emp_layout = QFormLayout()
        emp_layout.setContentsMargins(8, 4, 8, 8)
        emp_layout.setVerticalSpacing(4)
        
        self.emp_input = QLineEdit()
        self.emp_input.setPlaceholderText("Введите фамилию или ФИО")
        emp_layout.addRow(self.emp_input)

        action_btn_layout = QHBoxLayout()
        self.info_btn = QPushButton(QIcon(resource_path("resources/info.png")), "Info")
        self.cert_btn = QPushButton(QIcon(resource_path("resources/cert.png")), "Cert")
        self.disk_btn = QPushButton(QIcon(resource_path("resources/disk.png")), "Disk")
        
        for btn in [self.info_btn, self.cert_btn, self.disk_btn]:
            btn.setStyleSheet(element_style)
            action_btn_layout.addWidget(btn)
        
        emp_layout.addRow(action_btn_layout)
        emp_group.setLayout(emp_layout)
        main_layout.addWidget(emp_group)

        # Блок клиента
        client_group = QGroupBox("Клиент")
        client_group.setStyleSheet(element_style)
        client_layout = QFormLayout()
        client_layout.setContentsMargins(8, 4, 8, 8)
        client_layout.setVerticalSpacing(4)
        
        self.client_input = QLineEdit()
        self.client_input.setPlaceholderText("Введите название клиента")
        client_layout.addRow(self.client_input)
        
        client_btn_layout = QHBoxLayout()
        self.client_folder_btn = QPushButton(QIcon(resource_path("resources/folder.png")), "Client")
        self.all_certs_btn = QPushButton(QIcon(resource_path("resources/all.png")), "All Cert")
        
        for btn in [self.client_folder_btn, self.all_certs_btn]:
            btn.setStyleSheet(element_style)
            client_btn_layout.addWidget(btn)
        
        client_layout.addRow(client_btn_layout)
        client_group.setLayout(client_layout)
        main_layout.addWidget(client_group)

        # Настройки проверки срока действия
        self.expiry_check = QCheckBox("Проверять срок действия (15 месяцев)")
        self.expiry_check.setChecked(True)
        self.expiry_check.setStyleSheet("""
            QCheckBox {
                spacing: 5px;
                color: #e0e0e0;
            }
            QCheckBox::indicator {
                width: 16px;
                height: 16px;
            }
        """)
        main_layout.addWidget(self.expiry_check)

        # Блок подключения
        connect_group = QGroupBox("Прочее")
        connect_group.setStyleSheet(element_style)
        connect_layout = QHBoxLayout()
        connect_layout.setContentsMargins(8, 8, 8, 8)
        
        self.pc_input = QLineEdit()
        self.pc_input.setPlaceholderText("Имя компьютера")
        connect_layout.addWidget(self.pc_input, stretch=2)
        
        self.connect_btn = QPushButton(QIcon(resource_path("resources/connect.png")), "Connect")
        self.open_log_btn = QPushButton(QIcon(resource_path("resources/log.png")), "Log")
        self.laps_btn = QPushButton(QIcon(resource_path("resources/laps.png")), "LAPS")
        self.clear_btn = QPushButton(QIcon(resource_path("resources/clear.png")), "Clear")
        
        # Уменьшаем размер кнопок
        self.open_log_btn.setFixedWidth(40)
        self.laps_btn.setFixedWidth(60)
        self.clear_btn.setFixedWidth(80)
        
        for btn in [self.connect_btn, self.open_log_btn, self.laps_btn, self.clear_btn]:
            btn.setStyleSheet(element_style)
            connect_layout.addWidget(btn)
        
        connect_group.setLayout(connect_layout)
        main_layout.addWidget(connect_group)

        # Блок выполнения с кнопкой отмены
        self.execute_layout = QHBoxLayout()
        
        self.execute_btn = QPushButton(QIcon(resource_path("resources/execute.png")), "Выполнить")
        self.execute_btn.setStyleSheet("""
            QPushButton {
                background-color: #2a5699;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2e63b0;
            }
        """)
        self.execute_layout.addWidget(self.execute_btn, stretch=7)

        # Кнопки отмены и повтора
        self.undo_btn = QPushButton(QIcon(resource_path("resources/undo.png")), "")
        self.undo_btn.setStyleSheet("""
            QPushButton {
                background-color: #5a5a5a;
                min-width: 0;
                padding: 5px;
            }
            QPushButton:hover {
                background-color: #6a6a6a;
            }
            QPushButton:disabled {
                background-color: #3a3a3a;
            }
        """)
        self.undo_btn.setToolTip("Отменить последнее действие")
        self.undo_btn.setEnabled(False)
        
        self.redo_btn = QPushButton(QIcon(resource_path("resources/redo.png")), "")
        self.redo_btn.setStyleSheet("""
            QPushButton {
                background-color: #5a5a5a;
                min-width: 0;
                padding: 5px;
            }
            QPushButton:hover {
                background-color: #6a6a6a;
            }
            QPushButton:disabled {
                background-color: #3a3a3a;
            }
        """)
        self.redo_btn.setToolTip("Повторить последнее действие")
        self.redo_btn.setEnabled(False)
        
        self.execute_layout.addWidget(self.undo_btn, stretch=1)
        self.execute_layout.addWidget(self.redo_btn, stretch=1)
        
        main_layout.addLayout(self.execute_layout)

        # Лог-панель
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setStyleSheet("""
            QTextEdit {
                background-color: #252525;
                border: 1px solid #3a3a3a;
                font-family: 'Consolas';
                font-size: 12px;
                padding: 8px;
            }
        """)
        main_layout.addWidget(self.log_output)

        self.setLayout(main_layout)
        self.setup_completer()
    
    def setup_completer(self):
        """Настройка автодополнения для поля сотрудника"""
        self.completer_model = QStringListModel()
        self.completer = QCompleter()
        self.completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.completer.setModel(self.completer_model)
        self.emp_input.setCompleter(self.completer)
        self.emp_input.textChanged.connect(self.update_completer)

    def update_completer(self, text):
        """Обновление списка автодополнения с поиском только по началу фамилии"""
        if len(text.strip()) < 2:
            return
            
        # Ищем только по началу фамилии
        matches = [emp for emp in self.employee_manager.employees 
                   if emp.get('Фамилия', '').lower().startswith(text.strip().lower())]
        
        suggestions = []
        
        for emp in matches:
            last_name = emp.get('Фамилия', '')
            io = emp.get('ИО', '')
            if last_name and io:
                suggestions.append(f"{last_name} {io}")
        
        self.completer_model.setStringList(sorted(list(set(suggestions))))

    def setup_connections(self):
        """Настройка всех сигналов и слотов"""
        # Подключение кнопок сотрудника
        self.info_btn.clicked.connect(self.show_employee_info)
        self.cert_btn.clicked.connect(self.open_crypto_folder)
        self.disk_btn.clicked.connect(self.open_disk)
        
        # Подключение кнопок клиента
        self.client_folder_btn.clicked.connect(self.open_client_folder)
        self.all_certs_btn.clicked.connect(self.open_all_certs)
        
        # Подключение кнопок подключения
        self.connect_btn.clicked.connect(self.handle_connect)
        self.open_log_btn.clicked.connect(self.open_log_file)
        self.clear_btn.clicked.connect(self.clear_fields)
        self.laps_btn.clicked.connect(self.get_laps_password) 
        
        # Основные действия
        self.execute_btn.clicked.connect(self.execute_task)
        self.undo_btn.clicked.connect(self.undo_last_action)
        self.redo_btn.clicked.connect(self.redo_last_action)
        
        # Обработка нажатия Enter в полях ввода
        self.emp_input.returnPressed.connect(self.execute_btn.click)
        self.client_input.returnPressed.connect(self.execute_btn.click)
        self.pc_input.returnPressed.connect(self.execute_btn.click)
        
        # Сигнал для логирования
        self.log_signal.connect(self.log_message)
        
    def show_employee_info(self):
        """Отображение информации о сотруднике"""
        name_input = self.emp_input.text().strip()
        if not name_input:
            self.log_message("Ошибка: Введите фамилию сотрудника")
            return
            
        matches = self.employee_manager.search(name_input)
        if not matches:
            self.log_message(f"Сотрудник '{name_input}' не найден")
            return
            
        if len(matches) > 1:
            self.show_employee_selection(matches)
        else:
            employee = matches[0]
            self.emp_input.setText(f"{employee['Фамилия']} {employee['ИО']}")
            self.display_employee_info(employee)
        
    def log_search(self, query: str):
        """Логирование поиска только в интерфейс (без записи в файл)"""
        self.log_message(f"Поиск: {query}", "debug")

    def display_employee_info(self, employee: Dict):
        """Унифицированный формат информации о сотруднике"""
        info = (
            f"{employee.get('Фамилия', '')} {employee.get('ИО', '')} | "
            f"ВН: {employee.get('ВН', 'N/A')} | "
            f"Каб: {employee.get('Каб.', 'N/A')} | "
            f"ПК: {employee.get('ПК', 'N/A')} | "
            f"User: {employee.get('Username', 'N/A')}"
        )
        self.log_message(info)

    def show_employee_selection(self, employees: List[Dict], callback: Optional[callable] = None):
        """Показ выбора сотрудника с обработкой выбора"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Выбор сотрудника")
        dialog.setMinimumWidth(600)  # Устанавливаем минимальную ширину
        dialog.setStyleSheet("""
            QDialog {
                background-color: #252525;
                color: #e0e0e0;
            }
            QListWidget {
                background-color: #2d2d2d;
                color: #e0e0e0;
                border: 1px solid #3a3a3a;
                font-size: 12px;
            }
            QLabel {
                font-size: 12px;
            }
        """)
        
        layout = QVBoxLayout()
        label = QLabel("Найдено несколько совпадений:")
        layout.addWidget(label)
        
        list_widget = QListWidget()
        list_widget.setUniformItemSizes(True)  # Для оптимизации
        
        # Сортируем и добавляем только сотрудников, у которых фамилия начинается с введенного текста
        input_text = self.emp_input.text().strip().lower()
        filtered_employees = [
            emp for emp in employees 
            if emp.get('Фамилия', '').lower().startswith(input_text)
        ]
        
        # Если после фильтрации ничего не осталось, показываем всех
        employees_to_show = filtered_employees if filtered_employees else employees
        
        for emp in sorted(employees_to_show, key=lambda x: x.get('Фамилия', '')):
            item_text = (
                f"{emp.get('Фамилия', '')} {emp.get('ИО', '')} | "
                f"ВН: {emp.get('ВН', 'N/A')} | "
                f"Каб: {emp.get('Каб.', 'N/A')} | "
                f"ПК: {emp.get('ПК', 'N/A')}"
            )
            item = QListWidgetItem(item_text)
            list_widget.addItem(item)
        
        # Автоматически подбираем ширину по содержимому
        list_widget.setSizeAdjustPolicy(QListWidget.AdjustToContents)
        layout.addWidget(list_widget)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)
        
        dialog.setLayout(layout)
        
        if dialog.exec_() == QDialog.Accepted:
            selected_items = list_widget.selectedItems()
            if selected_items:
                index = list_widget.row(selected_items[0])
                selected_employee = employees_to_show[index]
                self.emp_input.setText(f"{selected_employee['Фамилия']} {selected_employee['ИО']}")
                self.display_employee_info(selected_employee)
                if callback:
                    callback(selected_employee)
                    
    def on_employee_selected(self, employee: Dict):
        """Обработка выбора сотрудника"""
        self.display_employee_info(employee)
        self.emp_input.setText(f"{employee['Фамилия']} {employee['ИО']}")


    def open_crypto_folder(self):
        """Открытие папки CryptoPro с четкой логикой"""
        surname = self.emp_input.text().strip()
        if not surname:
            self.log_message("Ошибка: Введите фамилию сотрудника", "error")
            return
        
        matches = self.employee_manager.search(surname)
        if not matches:
            self.log_message(f"Сотрудник '{surname}' не найден", "error")
            return
        
        if len(matches) > 1:
            self.show_employee_selection(matches, callback=self._open_crypto_for_employee)
        else:
            self._open_crypto_for_employee(matches[0])

    def _open_crypto_for_employee(self, employee: dict):
        """Внутренний метод с улучшенной обработкой ошибок"""
        pc_name = employee.get('ПК', '')
        username = employee.get('Username', '')
        
        if not pc_name or not username:
            self.log_message("Недостаточно данных для открытия папки (отсутствует ПК или логин)", "error")
            return
        
        path = self.employee_manager.get_crypto_path(pc_name, username)
        try:
            if path.exists():
                os.startfile(str(path))
                self.log_message(f"Открыта папка CryptoPro: {path}", "system")
            else:
                self.log_message(f"Папка CryptoPro не найдена: {path}", "error")
                # Предложение создать папку?
        except Exception as e:
            self.log_message(f"Ошибка доступа: {str(e)}", "error")

    def open_disk(self):
        """Открытие диска с каскадным выбором"""
        surname = self.emp_input.text().strip()
        pc_name = self.pc_input.text().strip()
        
        if not surname and not pc_name:
            self.log_message("Ошибка: Введите фамилию или имя ПК", "error")
            return
        
        if surname and pc_name:
            # Первый выбор: по сотруднику или по ПК
            self.show_disk_choice_dialog(surname, pc_name)
        elif surname:
            # Только сотрудник - сразу ищем сотрудников
            self._open_disk_by_employee(surname)
        else:
            # Только ПК - сразу открываем
            self._open_disk_by_pc(pc_name)

    def show_disk_choice_dialog(self, surname: str, pc_name: str):
        """Диалог первичного выбора (сотрудник или ПК)"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Выбор действия")
        dialog.setStyleSheet("""
            QDialog {
                background-color: #252525;
                color: #e0e0e0;
                padding: 10px;
            }
            QPushButton {
                min-width: 300px;
                padding: 8px;
                margin: 5px 0;
                text-align: left;
            }
        """)
        
        layout = QVBoxLayout()
        label = QLabel("Обнаружены оба параметра. Выберите действие:")
        layout.addWidget(label)
        
        # Кнопка открытия по сотруднику
        btn_emp = QPushButton(f"По сотруднику: {surname}")
        btn_emp.clicked.connect(lambda: (dialog.accept(), self._open_disk_by_employee(surname)))
        layout.addWidget(btn_emp)
        
        # Кнопка открытия по ПК
        btn_pc = QPushButton(f"По имени ПК: {pc_name}")
        btn_pc.clicked.connect(lambda: (dialog.accept(), self._open_disk_by_pc(pc_name)))
        layout.addWidget(btn_pc)
        
        # Кнопка отмены
        btn_cancel = QPushButton("Отмена")
        btn_cancel.clicked.connect(dialog.reject)
        layout.addWidget(btn_cancel)
        
        dialog.setLayout(layout)
        dialog.exec_()

    def _open_disk_by_employee(self, surname: str):
        """Открытие диска через выбор сотрудника"""
        matches = self.employee_manager.search(surname)
        if not matches:
            self.log_message(f"Сотрудник '{surname}' не найден", "error")
            return
            
        if len(matches) > 1:
            self.show_employee_selection(
                matches, 
                callback=lambda emp: self._open_disk_for_employee(emp)
            )
        else:
            self._open_disk_for_employee(matches[0])

    def _open_disk_for_employee(self, employee: dict):
        """Финальное открытие диска для выбранного сотрудника"""
        pc_name = employee.get('ПК', '')
        if not pc_name:
            self.log_message("У сотрудника не указан ПК", "error")
            return
            
        path = Path(f"\\\\{pc_name}\\c$")
        try:
            if path.exists():
                os.startfile(str(path))
                self.log_message(f"Открыт диск: {path}", "system")
            else:
                self.log_message(f"Не удалось открыть диск: {path}", "error")
        except Exception as e:
            self.log_message(f"Ошибка доступа: {str(e)}", "error")

    def _open_disk_by_pc(self, pc_name: str):
        """Непосредственное открытие диска по имени ПК"""
        path = Path(f"\\\\{pc_name}\\c$")
        try:
            if path.exists():
                os.startfile(str(path))
                self.log_message(f"Открыт диск: {path}", "system")
            else:
                self.log_message(f"Не удалось открыть диск: {path}", "error")
        except Exception as e:
            self.log_message(f"Ошибка доступа: {str(e)}", "error")
        
    def open_client_folder(self):
        """Открытие папки клиента"""
        client_name = self.client_input.text().strip()
        if not client_name:
            self.log_message("Ошибка: Введите название клиента")
            return
            
        clients = self.cert_manager.find_clients(client_name)
        if not clients:
            self.log_message(f"Клиент '{client_name}' не найден")
            return
            
        if len(clients) == 1:
            os.startfile(str(clients[0]))
            self.log_message(f"Открыта папка клиента: {clients[0].name}")
        else:
            self.show_client_selection(
                clients,
                callback=lambda client: (os.startfile(str(client))), 
                allow_cancel=True
            )

    def show_client_selection(self, clients: List[Path], callback: callable, allow_cancel: bool = True):
        """Показ выбора клиента с возможностью отмены"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Выбор клиента")
        dialog.setStyleSheet("""
            QDialog {
                background-color: #252525;
                color: #e0e0e0;
            }
            QListWidget {
                background-color: #2d2d2d;
                color: #e0e0e0;
                border: 1px solid #3a3a3a;
            }
        """)
        dialog.setMaximumWidth(600)
        
        layout = QVBoxLayout()
        label = QLabel("Найдено несколько клиентов:")
        layout.addWidget(label)
        
        list_widget = QListWidget()
        
        for client in sorted(clients, key=lambda x: x.name):
            list_widget.addItem(client.name)
        
        layout.addWidget(list_widget)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | (QDialogButtonBox.Cancel if allow_cancel else QDialogButtonBox.Ok))
        button_box.accepted.connect(dialog.accept)
        if allow_cancel:
            button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)
        
        dialog.setLayout(layout)
        
        if dialog.exec_() == QDialog.Accepted:
            selected_items = list_widget.selectedItems()
            if selected_items:
                index = list_widget.row(selected_items[0])
                callback(clients[index])
                
    def open_all_certs(self):
        """Открытие папки со всеми сертификатами"""
        try:
            os.startfile(str(NETWORK_FOLDER))
            self.log_message("Открыта папка со всеми сертификатами")
        except Exception as e:
            self.log_message(f"Ошибка открытия папки: {str(e)}")

    def handle_connect(self):
        """Обработка подключения к ПК с исправленной логикой"""
        surname = self.emp_input.text().strip()
        pc_name = self.pc_input.text().strip()
        
        if not surname and not pc_name:
            self.log_message("Ошибка: Введите фамилию или имя ПК", "error")
            return
        
        if surname and pc_name:
            # Показываем выбор действия
            self.show_dual_input_dialog(
                surname, 
                pc_name,
                action1=lambda: self.connect_by_name(surname),
                action2=lambda: self.connect_by_pc(pc_name),
                title1=f"Подключиться по сотруднику: {surname}",
                title2=f"Подключиться по ПК: {pc_name}"
            )
        elif surname:
            self.connect_by_name(surname)
        else:
            self.connect_by_pc(pc_name)

    def show_dual_input_dialog(self, surname: str, pc_name: str, action1: callable, action2: callable, 
                              title1: str, title2: str):
        """Универсальный диалог для выбора между двумя действиями"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Выбор действия")
        dialog.setStyleSheet("""
            QDialog {
                background-color: #252525;
                color: #e0e0e0;
                padding: 10px;
            }
            QPushButton {
                min-width: 300px;
                padding: 8px;
                margin: 5px 0;
                text-align: left;
            }
        """)
        
        layout = QVBoxLayout()
        msg = QLabel("Обнаружены оба параметра. Выберите действие:")
        layout.addWidget(msg)
        
        # Кнопка первого действия
        btn1 = QPushButton(title1)
        btn1.clicked.connect(lambda: (dialog.accept(), action1()))
        layout.addWidget(btn1)
        
        # Кнопка второго действия
        btn2 = QPushButton(title2)
        btn2.clicked.connect(lambda: (dialog.accept(), action2()))
        layout.addWidget(btn2)
        
        # Кнопка отмены
        btn_cancel = QPushButton("Отмена")
        btn_cancel.clicked.connect(dialog.reject)
        layout.addWidget(btn_cancel)
        
        dialog.setLayout(layout)
        dialog.exec_()

    def show_connection_choice(self, surname: str, pc_name: str, for_laps=False):
        """Показ выбора: по сотруднику или по ПК"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Выбор подключения")
        dialog.setStyleSheet("""
            QDialog {
                background-color: #252525;
                color: #e0e0e0;
            }
            QLabel {
                margin-bottom: 10px;
            }
            QPushButton {
                min-width: 200px;
                padding: 5px;
                margin: 5px 0;
                text-align: left;
            }
        """)
        
        layout = QVBoxLayout()
        
        # Информация о возможных вариантах
        info_label = QLabel("Обнаружены оба параметра - ФИО и имя ПК. Выберите способ подключения:")
        layout.addWidget(info_label)
        
        # Получаем информацию о сотруднике, если есть
        matches = self.employee_manager.search(surname)
        if matches:
            employee = matches[0]
            emp_pc = employee.get('ПК', '')
            if emp_pc and emp_pc.lower() != pc_name.lower():
                warning_label = QLabel(
                    f"⚠ Внимание: ПК сотрудника ({emp_pc}) не совпадает с введенным ПК ({pc_name})"
                )
                warning_label.setStyleSheet("color: #FFD700;")
                layout.addWidget(warning_label)
        
        # Кнопка подключения по сотруднику
        btn_by_emp = QPushButton(f"По сотруднику: {surname}")
        btn_by_emp.setToolTip("Будет использован ПК, указанный в базе для этого сотрудника")
        if for_laps:
            btn_by_emp.clicked.connect(lambda: (dialog.accept(), self.get_laps_by_name(surname)))
        else:
            btn_by_emp.clicked.connect(lambda: (dialog.accept(), self.connect_by_name(surname)))
        layout.addWidget(btn_by_emp)
        
        # Кнопка подключения по ПК
        btn_by_pc = QPushButton(f"По имени ПК: {pc_name}")
        btn_by_pc.setToolTip("Будет использовано прямое подключение по указанному имени ПК")
        if for_laps:
            btn_by_pc.clicked.connect(lambda: (dialog.accept(), self.get_laps_by_pc(pc_name)))
        else:
            btn_by_pc.clicked.connect(lambda: (dialog.accept(), self.connect_by_pc(pc_name)))
        layout.addWidget(btn_by_pc)
        
        # Кнопка отмены
        cancel_btn = QPushButton("Отмена")
        cancel_btn.clicked.connect(dialog.reject)
        layout.addWidget(cancel_btn)
        
        dialog.setLayout(layout)
        dialog.exec_()
        
    def connect_by_name(self, surname: str):
        """Подключение по имени сотрудника с гарантированным выполнением"""
        matches = self.employee_manager.search(surname)
        if not matches:
            self.log_message(f"Сотрудник '{surname}' не найден", "error")
            return
            
        if len(matches) > 1:
            # Передаем callback, который выполнит подключение после выбора
            self.show_employee_selection(
                matches,
                callback=lambda emp: self._perform_connection(emp)
            )
        else:
            self._perform_connection(matches[0])

    def connect_by_pc(self, pc_name: str):
        """Подключение по имени ПК с базовой валидацией"""
        if not pc_name.strip():
            self.log_message("Ошибка: Введите имя ПК", "error")
            return
        
        # Обновляем поле ПК (на случай, если были лишние пробелы)
        self.pc_input.setText(pc_name.strip())
        
        # Логируем попытку подключения
        self.log_message(f"Попытка подключения к ПК: {pc_name}", "system")
        
        # Вызываем подтверждение подключения без информации о сотруднике
        self.confirm_connection(pc_name.strip(), None)

    def _perform_connection(self, employee: dict):
        """Фактическое выполнение подключения"""
        pc_name = employee.get('ПК', '')
        if not pc_name:
            self.log_message("У сотрудника не указан ПК", "error")
            return
        
        # Обновляем поле сотрудника
        self.emp_input.setText(f"{employee['Фамилия']} {employee['ИО']}")
        self.display_employee_info(employee)
        
        # Выполняем подключение
        self.confirm_connection(pc_name, employee)
    
    def confirm_connection(self, pc_name: str, employee: Optional[dict] = None):
        """Подтверждение подключения с полной информацией"""
        if employee:
            info = (
                f"{employee.get('Фамилия', '')} {employee.get('ИО', '')} | "
                f"ВН: {employee.get('ВН', 'N/A')} | "
                f"Каб: {employee.get('Каб.', 'N/A')} | "
                f"ПК: {employee.get('ПК', 'N/A')} | "
                f"User: {employee.get('Username', 'N/A')}"
            )
            message = f"Подключиться к сотруднику:\n{info}\n\nКомпьютер: {pc_name}?"
        else:
            message = f"Подключиться к компьютеру: {pc_name}?"
        
        reply = QMessageBox.question(
            self,
            "Подтверждение",
            message,
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            if self.cert_manager.connect_to_pc(pc_name):
                self.log_message(f"Успешное подключение к {pc_name}", "success")
                self.log_connection(pc_name, employee)
            else:
                self.log_message(f"Ошибка подключения к {pc_name}", "error")

    def log_connection(self, pc_name: str, employee: Optional[dict] = None):
        """Логирование подключения только в интерфейс (без записи в файл)"""
        if employee:
            log_msg = f"Подключение к {pc_name} ({employee.get('Фамилия', '')} {employee.get('ИО', '')})"
        else:
            log_msg = f"Подключение к {pc_name}"
        
        self.log_message(log_msg, "system")

    def clear_fields(self):
        """Очистка полей ввода"""
        self.emp_input.clear()
        self.client_input.clear()
        self.pc_input.clear()
        self.log_message("Поля очищены")

    def execute_task(self):
        """Исправленный метод выполнения задачи"""
        surname = self.emp_input.text().strip()
        client_name = self.client_input.text().strip()
        
        if not surname or not client_name:
            self.log_message("Ошибка: Заполните все обязательные поля", "error")
            return
            
        # Поиск сотрудника
        matches = self.employee_manager.search(surname)
        if not matches:
            self.log_message(f"Ошибка: Сотрудник '{surname}' не найден", "error")
            return
            
        if len(matches) > 1:
            self.show_employee_selection(matches, callback=lambda emp: self.process_client_selection(emp))
            return
            
        # Поиск клиента
        clients = self.cert_manager.find_clients(client_name)
        if not clients:
            self.log_message(f"Ошибка: Клиент '{client_name}' не найден", "error")
            return
            
        if len(clients) > 1:
            self.show_client_selection(
                clients,
                callback=lambda client: self.copy_certificates(client, matches[0])
            )
            return
            
        self.copy_certificates(clients[0], matches[0])
    
    def process_client_selection(self, employee):
        """Обработка выбора клиента для копирования"""
        client_name = self.client_input.text().strip()
        clients = self.cert_manager.find_clients(client_name)
        
        if not clients:
            self.log_message(f"Ошибка: Клиент '{client_name}' не найден")
            return
            
        if len(clients) == 1:
            self.copy_certificates(clients[0], employee)
        else:
            self.show_client_selection(
                clients,
                callback=lambda client: self.copy_certificates(client, employee)
            )
        
    def copy_certificates(self, client_path: Path, employee: Dict):
        """Копирование сертификатов с сохранением информации для отмены/повтора"""
        check_expiry = self.expiry_check.isChecked()
        certs = self.cert_manager.get_certificates(client_path)
        
        if not certs:
            self.log_message(f"У клиента {client_path.name} нет сертификатов")
            return
            
        if check_expiry:
            valid_certs = [c for c in certs if c["status"] == "valid"]
            if not valid_certs:
                self.log_message(f"Нет действительных сертификатов у клиента {client_path.name}")
                return
            cert_to_copy = valid_certs[0]
        else:
            cert_to_copy = certs[0]
            
        pc_name = employee.get("ПК")
        username = employee.get("Username")
        if not pc_name or not username:
            self.log_message("Недостаточно данных о сотруднике для копирования")
            return
            
        dest_path = self.employee_manager.get_crypto_path(pc_name, username) / cert_to_copy["name"]
        success, message = self.cert_manager.copy_certificate(cert_to_copy["path"], dest_path)
        
        if success:
            # Сохраняем информацию для отмены/повтора
            self.last_action = {
                'type': 'copy',
                'source': cert_to_copy["path"],
                'destination': dest_path,
                'employee': employee,
                'client': client_path.name
            }
            self.undo_btn.setEnabled(True)
            self.redo_btn.setEnabled(False)
            
            self.log_message(f"Сертификат {client_path.name} ({cert_to_copy['name']}) скопирован")
            
            # Информация о сотруднике
            self.display_employee_info(employee)
        else:
            self.log_message(f"Ошибка: {message}")

    def open_log_file(self):
        """Открытие лог-файла"""
        log_dir = LOG_FOLDER
        today = datetime.now().strftime("%d-%m-%Y")
        log_path = log_dir / f"{today}.log"
        
        try:
            if not log_dir.exists():
                self.log_message("Ошибка: Папка логов недоступна")
                return
                
            if not log_path.exists():
                self.log_message(f"Лог-файл за {today} не найден")
                return
                
            os.startfile(str(log_path))
            self.log_message(f"Открыт лог-файл: {log_path.name}")
        except Exception as e:
            self.log_message(f"Ошибка открытия лога: {str(e)}")

    def log_message(self, message: str, msg_type: Optional[str] = None):
        """
        Улучшенное логирование с чётким определением типов сообщений и цветовым оформлением.
        
        Параметры:
            message (str): Текст сообщения для логирования
            msg_type (Optional[str]): Тип сообщения (определяет цвет и иконку):
                - 'error': Ошибки и критические проблемы
                - 'success': Успешные операции
                - 'info': Информационные сообщения
                - 'warning': Предупреждения
                - 'action': Действия пользователя (отмена/повтор)
                - 'system': Системные сообщения
                - 'debug': Отладочная информация
                - None: Автоматическое определение
        """

        timestamp = datetime.now().strftime("%H:%M:%S")
        
        if msg_type is None:
            msg_type = self._detect_message_type(message)
        
        style = self._get_message_style(msg_type)
        
        # Формируем HTML для сообщения
        html_message = self._format_html_message(timestamp, message, style)
        
        # Сохраняем состояние прокрутки
        scrollbar = self.log_output.verticalScrollBar()
        at_bottom = scrollbar.value() == scrollbar.maximum()
        
        # Вставляем сообщение в лог
        cursor = self.log_output.textCursor()
        cursor.movePosition(QTextCursor.End)
        
        # Вставляем HTML и принудительный перенос строки
        cursor.insertHtml(html_message)
        cursor.insertBlock()  # Это добавит новый абзац/строку
        
        # Прокручиваем вниз если пользователь был в конце лога
        if at_bottom:
            scrollbar.setValue(scrollbar.maximum())
        
        # Ограничение размера лога
        self._trim_log_content()

    def _detect_message_type(self, message: str) -> str:
        """Определяет тип сообщения по его содержанию"""
        msg_lower = message.lower()
        
        # Ошибки
        if any(word in msg_lower for word in [
            'ошибка', 'не найд', 'не удалось', 'не указан', 
            'недоступн', 'отказано', 'провал', 'failed', 'error'
        ]):
            return 'error'
        
        # Успешные операции
        if any(word in msg_lower for word in [
            'успешн', 'скопирован', 'подключени', 'готов', 'установлен',
            'success', 'copied', 'connected', 'done'
        ]):
            return 'success'
        
        # Информация о сотруднике
        if any(sep in message for sep in ['|', 'ВН:', 'Каб:', 'ПК:', 'User:']):
            return 'info'
        
        # Предупреждения
        if any(word in msg_lower for word in [
            'внимание', 'предупрежден', 'warning', 'проверьте', 'убедитесь'
        ]):
            return 'warning'
        
        # Действия пользователя
        if any(word in msg_lower for word in [
            'отмен', 'повтор', 'действие', 'action', 'undo', 'redo'
        ]):
            return 'action'
        
        # Системные сообщения
        if any(word in msg_lower for word in [
            'открыт', 'папк', 'диск', 'файл', 'folder', 'file', 'disk'
        ]):
            return 'system'
        
        # По умолчанию
        return 'default'

    def _get_message_style(self, msg_type: str) -> Dict:
        """Возвращает стиль оформления для типа сообщения"""
        styles = {
            'error': {
                'color': '#FF6B6B',  # Красный
                'icon': '✗',
                'bg': '#3A1E1E',
                'bold': True
            },
            'success': {
                'color': '#76FF7A',  # Зелёный
                'icon': '✓',
                'bg': '#1E3A1E',
                'bold': True
            },
            'info': {
                'color': '#4DA6FF',  # Синий
                'icon': 'ℹ',
                'bg': '#1E2A3A',
                'bold': False
            },
            'warning': {
                'color': '#FFD700',  # Жёлтый
                'icon': '⚠',
                'bg': '#3A321E',
                'bold': True
            },
            'action': {
                'color': '#FFA500',  # Оранжевый
                'icon': '↺',
                'bg': '#3A2A1E',
                'bold': True
            },
            'system': {
                'color': '#BA8CFF',  # Фиолетовый
                'icon': '⚙',
                'bg': '#2A1E3A',
                'bold': False
            },
            'debug': {
                'color': '#AAAAAA',  # Серый
                'icon': '⚐',
                'bg': '#2A2A2A',
                'bold': False
            },
            'default': {
                'color': '#E0E0E0',  # Светло-серый
                'icon': '',
                'bg': '#252525',
                'bold': False
            }
        }
        return styles.get(msg_type, styles['default'])

    def _format_html_message(self, timestamp: str, message: str, style: Dict) -> str:
        """Форматирование сообщения без фона строки"""
        icon = f'<span style="color:{style["color"]}">{style["icon"]}</span> ' if style["icon"] else ''
        
        return (
            f'<div style="margin: 2px 0; padding: 2px 0;">'  # Убрали background-color
            f'<span style="color: #AAAAAA; font-size: 10px;">[{timestamp}]</span> '
            f'{icon}'
            f'<span style="color: {style["color"]}; font-weight: {"bold" if style["bold"] else "normal"}">{message}</span>'
            f'</div>'
        )

    def _trim_log_content(self):
        """Ограничивает размер лога, удаляя старые сообщения"""
        max_lines = 1000
        doc = self.log_output.document()
        
        if doc.lineCount() > max_lines:
            cursor = QTextCursor(doc)
            cursor.movePosition(QTextCursor.Start)
            
            # Удаляем блоками по 50 строк для производительности
            lines_to_remove = doc.lineCount() - max_lines + 50
            cursor.movePosition(QTextCursor.Down, QTextCursor.KeepAnchor, lines_to_remove)
            cursor.removeSelectedText()

    def undo_last_action(self):
        """Отмена последнего действия с возможностью повтора"""
        if not hasattr(self, 'last_action') or not self.last_action:
            self.log_message("Нет действий для отмены")
            return
            
        if self.last_action['type'] == 'copy':
            try:
                if self.last_action['destination'].exists():
                    # Сохраняем действие для возможного повтора
                    self.redo_action = self.last_action.copy()
                    shutil.rmtree(str(self.last_action['destination']))
                    self.log_message(f"Отменено: сертификат {self.last_action['destination'].name} удален")
                    self.undo_btn.setEnabled(False)
                    self.redo_btn.setEnabled(True)
                    self.last_action = None
            except Exception as e:
                self.log_message(f"Ошибка отмены: {str(e)}")

    def redo_last_action(self):
        """Повтор последнего отмененного действия"""
        if not hasattr(self, 'redo_action') or not self.redo_action:
            self.log_message("Нет действий для повтора")
            return
            
        try:
            success, message = self.cert_manager.copy_certificate(
                self.redo_action['source'],
                self.redo_action['destination']
            )
            
            if success:
                self.log_message(f"Повторено: сертификат {self.redo_action['destination'].name} скопирован")
                self.last_action = self.redo_action
                self.undo_btn.setEnabled(True)
                self.redo_btn.setEnabled(False)
                self.redo_action = None
            else:
                self.log_message(f"Ошибка повтора: {message}")
        except Exception as e:
            self.log_message(f"Ошибка повтора: {str(e)}")
            
    def get_laps_password(self):
        """Получение пароля LAPS с умной логикой выбора"""
        surname = self.emp_input.text().strip()
        pc_name = self.pc_input.text().strip()
        
        if not surname and not pc_name:
            self.log_message("Ошибка: Введите фамилию или имя ПК", "error")
            return
        
        def execute_for_pc(pc: str):
            password = self.get_laps_password_from_ad(pc)
            if password:
                self.show_laps_dialog(pc, password, None)
            else:
                self.log_message(f"Не удалось получить пароль LAPS для {pc}", "error")
        
        if surname and pc_name:
            self.show_dual_input_dialog(
                surname,
                pc_name,
                action1=lambda: self.get_laps_by_name(surname),
                action2=lambda: execute_for_pc(pc_name),
                title1=f"Получить LAPS для сотрудника {surname}",
                title2=f"Получить LAPS для ПК {pc_name}"
            )
        elif surname:
            self.get_laps_by_name(surname)
        else:
            execute_for_pc(pc_name)

    def get_laps_by_name(self, surname: str):
        """Получение пароля LAPS по имени сотрудника"""
        matches = self.employee_manager.search(surname)
        if not matches:
            self.log_message(f"Сотрудник '{surname}' не найден")
            return
            
        if len(matches) > 1:
            self.show_employee_selection(matches, callback=lambda emp: self.fetch_and_show_laps(emp.get('ПК', ''), emp))
            return
            
        employee = matches[0]
        pc_name = employee.get('ПК', '')
        
        if not pc_name:
            self.log_message("У сотрудника не указан ПК")
            return
            
        self.fetch_and_show_laps(pc_name, employee)

    def get_laps_by_pc(self, pc_name: str):
        """Получение пароля LAPS по имени ПК"""
        # Поиск сотрудника по ПК
        employee = None
        for emp in self.employee_manager.employees:
            if emp.get('ПК', '') == pc_name:
                employee = emp
                break
                
        self.fetch_and_show_laps(pc_name, employee)

    def fetch_and_show_laps(self, pc_name: str, employee: dict):
        """Запрос пароля LAPS и отображение"""
        password = self.get_laps_password_from_ad(pc_name)
        if password is None:
            self.log_message(f"Не удалось получить пароль LAPS для {pc_name}")
            return
            
        # Отображаем пароль в диалоговом окне
        self.show_laps_dialog(pc_name, password, employee)

    def get_laps_password_from_ad(self, computer_name: str) -> str:
        """Получение пароля LAPS из Active Directory с использованием текущего контекста безопасности"""
        try:
            # Формируем PowerShell скрипт
            ps_script = f'''
            $computerName = "{computer_name}"
            $password = $null
            
            # Проверяем наличие модуля AdmPwd.PS
            if (Get-Module -ListAvailable -Name AdmPwd.PS) {{
                Import-Module AdmPwd.PS
                $result = Get-AdmPwdPassword -ComputerName $computerName -ErrorAction SilentlyContinue
                if ($result) {{ $password = $result.Password }}
            }}
            
            if (-not $password) {{
                # Попробуем через ADSI
                $searcher = [ADSISearcher]"(cn=$computerName)"
                $searcher.PropertiesToLoad.Add("ms-Mcs-AdmPwd") | Out-Null
                $result = $searcher.FindOne()
                if ($result) {{ $password = $result.Properties["ms-mcs-admpwd"][0] }}
            }}
            
            if ($password) {{ Write-Output $password }}
            else {{ exit 1 }}
            '''
            
            # Запускаем PowerShell с текущими учетными данными пользователя
            result = subprocess.run(
                ["powershell", "-Command", ps_script],
                capture_output=True,
                text=True,
                encoding='utf-8',
                timeout=10
            )
            
            if result.returncode == 0:
                return result.stdout.strip()
            else:
                print(f"PowerShell error: {result.stderr}")
                return None
                
        except Exception as e:
            print(f"Ошибка получения LAPS: {e}")
            return None

    def is_laps_module_available(self):
        """Проверка доступности модуля LAPS"""
        try:
            import AdmPwd.PS
            return True
        except ImportError:
            return False

    def show_laps_dialog(self, pc_name: str, password: str, employee: dict):
        """Диалоговое окно с паролем"""
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Пароль LAPS для {pc_name}")
        layout = QVBoxLayout()
        
        # Отображаем информацию о сотруднике, если есть
        if employee:
            info_label = QLabel(
                f"Сотрудник: {employee.get('Фамилия', '')} {employee.get('ИО', '')}\n"
                f"ВН: {employee.get('ВН', 'N/A')} | Каб: {employee.get('Каб.', 'N/A')}\n"
                f"ПК: {employee.get('ПК', 'N/A')} | User: {employee.get('Username', 'N/A')}"
            )
            layout.addWidget(info_label)
        
        # Пароль
        password_label = QLabel("Пароль администратора:")
        layout.addWidget(password_label)
        
        password_edit = QLineEdit(password)
        password_edit.setReadOnly(True)
        layout.addWidget(password_edit)
        
        # Кнопки
        btn_layout = QHBoxLayout()
        copy_btn = QPushButton("Копировать")
        copy_btn.clicked.connect(lambda: self.copy_to_clipboard(password))
        btn_layout.addWidget(copy_btn)
        
        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(dialog.accept)
        btn_layout.addWidget(close_btn)
        
        layout.addLayout(btn_layout)
        dialog.setLayout(layout)
        dialog.exec_()

    def copy_to_clipboard(self, text: str):
        """Копирование текста в буфер обмена"""
        clipboard = QApplication.clipboard()
        clipboard.setText(text)
        self.log_message("Пароль скопирован в буфер обмена")

    def show_connection_choice(self, surname: str, pc_name: str, for_laps=False):
        """Показ выбора: по сотруднику или по ПК"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Выбор подключения")
        dialog.setStyleSheet("background-color: #252525; color: #e0e0e0;")
        
        layout = QVBoxLayout()
        
        # Получаем информацию о сотруднике
        employee_info = ""
        matches = self.employee_manager.search(surname)
        if matches:
            employee = matches[0]
            employee_info = (
                f"Сотрудник: {employee.get('Фамилия', '')} {employee.get('ИО', '')}\n"
                f"ВН: {employee.get('ВН', 'N/A')} | Каб: {employee.get('Каб.', 'N/A')}\n"
                f"ПК: {employee.get('ПК', 'N/A')} | User: {employee.get('Username', 'N/A')}"
            )
            info_label = QLabel(employee_info)
            layout.addWidget(info_label)
        
        label = QLabel("Выберите способ подключения:")
        layout.addWidget(label)
        
        btn1 = QPushButton(f"По сотруднику: {surname}")
        if for_laps:
            btn1.clicked.connect(lambda: (dialog.accept(), self.get_laps_by_name(surname)))
        else:
            btn1.clicked.connect(lambda: (dialog.accept(), self.connect_by_name(surname)))
        layout.addWidget(btn1)
        
        btn2 = QPushButton(f"По ПК: {pc_name}")
        if for_laps:
            btn2.clicked.connect(lambda: (dialog.accept(), self.get_laps_by_pc(pc_name)))
        else:
            btn2.clicked.connect(lambda: (dialog.accept(), self.connect_by_pc(pc_name)))
        layout.addWidget(btn2)
        
        cancel_btn = QPushButton("Отмена")
        cancel_btn.clicked.connect(dialog.reject)
        layout.addWidget(cancel_btn)
        
        dialog.setLayout(layout)
        
        # Рассчитываем оптимальную ширину
        font_metrics = dialog.fontMetrics()
        max_width = max(
            font_metrics.horizontalAdvance(f"По сотруднику: {surname}"),
            font_metrics.horizontalAdvance(f"По ПК: {pc_name}"),
            font_metrics.horizontalAdvance("Отмена")
        )
        if employee_info:
            max_width = max(max_width, font_metrics.horizontalAdvance(employee_info.split('\n')[0]))
            
        dialog.setMinimumWidth(min(max_width + 50, 800))
        dialog.adjustSize()
        dialog.exec_()