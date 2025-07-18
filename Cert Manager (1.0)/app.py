import sys
from pathlib import Path
from datetime import datetime
from PySide6.QtWidgets import QApplication
from ui.main_window import MainWindow
from config import NETWORK_FOLDER, ARCHIVE_FOLDER, EXCEL_FILE, CRYPTO_PRO_PATH

def check_log_folder() -> tuple[bool, str]:
    """Проверка доступности папки логов"""
    log_dir = Path(r"\\nas\Distrib\script\certificate\cert_log")
    today = datetime.now().strftime("%d-%m-%Y")
    log_file = log_dir / f"{today}.log"
    
    try:
        if not log_dir.exists():
            return False, "Папка логов недоступна"
        return True, f"Лог-файл {'существует' if log_file.exists() else 'не существует'}"
    except Exception as e:
        return False, f"Ошибка проверки логов: {str(e)}"

def verify_excel_columns(window: MainWindow) -> bool:
    """Проверка структуры файла сотрудников"""
    try:
        import pandas as pd
        df = pd.read_excel(EXCEL_FILE)
        required_columns = ['Фамилия', 'ИО', 'ПК', 'Username']
        missing = [col for col in required_columns if col not in df.columns]
        
        if missing:
            window.copy_tab.log_signal.emit(
                f"Ошибка: В файле отсутствуют столбцы: {', '.join(missing)}", 
                "error"
            )
            return False
        return True
    except Exception as e:
        window.copy_tab.log_signal.emit(
            f"Ошибка проверки файла: {str(e)}", 
            "error"
        )
        return False

def run_startup_tests(window: MainWindow):
    """Запуск тестов системы при старте"""
    tests = [
        ("Доступ к сетевой папке сертификатов", NETWORK_FOLDER.exists()),
        ("Доступ к архиву сертификатов", ARCHIVE_FOLDER.exists()),
        ("Доступ к файлу сотрудников", EXCEL_FILE.exists()),
        ("Доступ к Crypto Pro", CRYPTO_PRO_PATH.exists()),
    ]
    
    log_status, log_msg = check_log_folder()
    tests.append(("Проверка папки логов", log_status))
    
    window.copy_tab.log_signal.emit("=== Запуск тестов системы ===", "info")
    for name, result in tests:
        status = "Успех" if result else "Ошибка"
        msg_type = "success" if result else "error"
        window.copy_tab.log_signal.emit(f"{status}: {name}", msg_type)
    
    if log_status:
        window.copy_tab.log_signal.emit(f"Инфо: {log_msg}", "info")
    
    verify_excel_columns(window)
    window.copy_tab.log_signal.emit("=== Тесты завершены ===", "info")

def main():
    app = QApplication(sys.argv)
    
    # Применяем стили ко всему приложению
    app.setStyle("Fusion")  # Используем Fusion стиль как основу
    
    window = MainWindow()
    run_startup_tests(window)
    
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()