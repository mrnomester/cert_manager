import shutil
import subprocess
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Dict, Tuple
from config import NETWORK_FOLDER, ARCHIVE_FOLDER, CERT_EXPIRY_DAYS, CRYPTO_PRO_PATH
from concurrent.futures import ThreadPoolExecutor
from functools import lru_cache

class CertificateManager:
    @staticmethod
    @lru_cache(maxsize=100)
    def find_clients(query: str) -> List[Path]:
        query = query.lower()
        results = []
        
        def search_in_folder(folder: Path):
            for item in folder.iterdir():
                if item.is_dir() and query in item.name.lower():
                    results.append(item)
        
        with ThreadPoolExecutor() as executor:
            executor.submit(search_in_folder, NETWORK_FOLDER)
            executor.submit(search_in_folder, ARCHIVE_FOLDER)
            
        return results

    @staticmethod
    def get_certificates(client_path: Path) -> List[Dict]:
        if not client_path.exists():
            return []
            
        expiry_date = datetime.now() - timedelta(days=CERT_EXPIRY_DAYS)
        certs = []
        
        for cert_dir in client_path.iterdir():
            if cert_dir.is_dir():
                mod_time = datetime.fromtimestamp(cert_dir.stat().st_mtime)
                status = "valid" if mod_time >= expiry_date else "expired"
                certs.append({
                    "path": cert_dir,
                    "name": cert_dir.name,
                    "status": status,
                    "date": mod_time.strftime("%d.%m.%Y %H:%M"),
                    "full_path": str(cert_dir)
                })
                
        return sorted(certs, key=lambda x: x["path"].stat().st_mtime, reverse=True)

    @staticmethod
    def copy_certificate(src: Path, dest: Path) -> Tuple[bool, str]:
        try:
            if not src.exists():
                return False, f"Источник не существует: {src}"
            if dest.exists():
                return False, f"Целевая папка уже существует: {dest}"
            shutil.copytree(src, dest)
            return True, f"Сертификат скопирован в {dest}"
        except PermissionError as e:
            return False, f"Ошибка доступа: {str(e)}"
        except Exception as e:
            return False, f"Неизвестная ошибка: {str(e)}"

    @staticmethod
    def delete_certificate(cert_path: Path) -> Tuple[bool, str]:
        try:
            if not cert_path.exists():
                return False, "Сертификат не найден"
            if not cert_path.is_dir():
                return False, "Указанный путь не является папкой"
            shutil.rmtree(cert_path)
            return True, f"Сертификат {cert_path.name} удален"
        except PermissionError as e:
            return False, f"Нет прав на удаление: {str(e)}"
        except Exception as e:
            return False, f"Ошибка удаления: {str(e)}"

    @staticmethod
    def connect_to_pc(pc_name: str) -> bool:
        try:
            if not pc_name:
                return False
                
            command = f'"C:\\Program Files (x86)\\SolarWinds\\DameWare Remote Support\\dwrcc.exe" -c -m:{pc_name} -a:1'
            subprocess.Popen(command, shell=True)
            return True
        except Exception as e:
            print(f"Ошибка подключения: {e}")
            return False