#!/usr/bin/env python3
import os
import subprocess
from OpenSSL import crypto
from datetime import datetime
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
import time

class CertificateParserGUI:
    def __init__(self):
        self.max_files = 5000
        self.buffer_size = 1024 * 1024  # 1MB буфер
        self.max_workers = 4
        self.current_dir = ""
        self.progress = None

    def show_dialog(self, title, text, dialog_type='info', width=400):
        cmd = ['zenity', f'--{dialog_type}', f'--title={title}', 
               f'--text={text}', f'--width={width}']
        return subprocess.run(cmd, check=False)

    def get_directory(self):
        try:
            result = subprocess.run(
                ['zenity', '--file-selection', '--directory',
                 '--title=Выберите папку с сертификатами .cer',
                 '--width=600', '--height=400'],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
            return result.stdout.strip() if result.returncode == 0 else None
        except FileNotFoundError:
            self.show_dialog('Ошибка', 'Zenity не установлен!\nУстановите: sudo apt install zenity', 'error')
            return None

    def get_output_file(self):
        result = subprocess.run(
            ['zenity', '--file-selection', '--save',
             '--title=Сохранить отчет как...',
             '--filename=report.xlsx',
             '--file-filter=*.xlsx',
             '--width=600', '--height=400'],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        return result.stdout.strip() if result.returncode == 0 else None

    def start_progress(self, total_files):
        """Запуск прогресс-бара с отображением текущей директории"""
        self.progress = subprocess.Popen(
            ['zenity', '--progress',
             '--title=Анализ сертификатов',
             '--text=Начало обработки...\nТекущая директория: ',
             '--percentage=0',
             '--auto-close',
             '--no-cancel',
             '--width=500',
             '--height=200'],
            stdin=subprocess.PIPE,
            text=True
        )
        self.total_files = total_files
        self.processed = 0

    def update_progress(self, current_file):
        """Обновление прогресса с отображением текущего пути"""
        if not self.progress or self.progress.poll() is not None:
            return

        self.processed += 1
        percent = int((self.processed / self.total_files) * 100)
        current_dir = os.path.dirname(current_file)
        
        # Обновляем только если директория изменилась
        if current_dir != self.current_dir:
            self.current_dir = current_dir
            display_text = (
                f"Обработка файла: {os.path.basename(current_file)}\n"
                f"Текущая директория: {current_dir}"
            )
            self.progress.stdin.write(f"#{display_text}\n")
        
        self.progress.stdin.write(f"{percent}\n")
        self.progress.stdin.flush()

    def parse_certificate(self, cert_path):
        """Парсинг сертификата с обработкой ошибок"""
        try:
            with open(cert_path, 'rb', buffering=self.buffer_size) as f:
                cert = crypto.load_certificate(crypto.FILETYPE_ASN1, f.read())
                
                subject = cert.get_subject()
                fio = " ".join(
                    getattr(subject, field) 
                    for field in ['CN', 'SN', 'GN', 'givenName', 'surname'] 
                    if hasattr(subject, field))
                
                return {
                    'ФИО': fio.strip(),
                    'Дата создания': datetime.strptime(
                        cert.get_notBefore().decode('utf-8'), '%Y%m%d%H%M%SZ'),
                    'Дата окончания': datetime.strptime(
                        cert.get_notAfter().decode('utf-8'), '%Y%m%d%H%M%SZ'),
                    'Серийный номер': format(cert.get_serial_number(), 'X'),
                    'Путь до файла': cert_path
                }
        except Exception as e:
            print(f"Ошибка в {cert_path}: {str(e)}")
            return None

    def find_cert_files(self, directory):
        """Рекурсивный поиск .cer файлов"""
        cer_files = []
        for root, _, files in os.walk(directory):
            current_files = [
                os.path.join(root, f) 
                for f in files 
                if f.lower().endswith('.cer')
            ]
            cer_files.extend(current_files)
            
            # Обновляем отображение текущей директории
            if self.progress and self.progress.poll() is None:
                display_text = (
                    f"Сканирование: {root}\n"
                    f"Найдено файлов: {len(cer_files)}"
                )
                self.progress.stdin.write(f"#{display_text}\n")
                self.progress.stdin.write("0\n")  # Процент не меняем при сканировании
                self.progress.stdin.flush()
            
            if len(cer_files) >= self.max_files:
                break
        
        return cer_files[:self.max_files]

    def process_directory(self, directory):
        """Многопоточная обработка с визуализацией"""
        # Начинаем сканирование с отображением прогресса
        self.start_progress(1)  # Временное значение
        cer_files = self.find_cert_files(directory)
        
        if not cer_files:
            if self.progress:
                self.progress.stdin.close()
            self.show_dialog('Ошибка', 'Не найдено .cer файлов', 'error')
            return []
        
        # Перезапускаем прогресс-бар с реальным количеством файлов
        if self.progress:
            self.progress.stdin.close()
            self.progress.wait()
        
        self.start_progress(len(cer_files))
        results = []
        
        try:
            with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
                futures = {executor.submit(self.parse_certificate, path): path 
                          for path in cer_files}
                
                for future in as_completed(futures):
                    current_file = futures[future]
                    self.update_progress(current_file)
                    
                    if res := future.result():
                        results.append(res)
        
        finally:
            if self.progress and self.progress.poll() is None:
                self.progress.stdin.close()
                self.progress.wait()
        
        return results

    def save_to_excel(self, data, output_file):
        """Сохранение результатов в Excel"""
        if not data:
            self.show_dialog('Ошибка', 'Нет данных для сохранения', 'error')
            return False
        
        try:
            # Показываем прогресс сохранения
            with subprocess.Popen(
                ['zenity', '--progress',
                 '--title=Сохранение отчета',
                 '--text=Сохранение в Excel...',
                 '--percentage=0',
                 '--auto-close',
                 '--no-cancel',
                 '--width=400'],
                stdin=subprocess.PIPE,
                text=True
            ) as progress:
                
                df = pd.DataFrame(data)
                progress.stdin.write("#Форматирование дат...\n25\n")
                progress.stdin.flush()
                
                df['Дата создания'] = df['Дата создания'].dt.strftime('%d.%m.%Y %H:%M:%S')
                df['Дата окончания'] = df['Дата окончания'].dt.strftime('%d.%m.%Y %H:%M:%S')
                
                progress.stdin.write("#Запись в файл...\n50\n")
                progress.stdin.flush()
                
                with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                
                progress.stdin.write("#Готово!\n100\n")
                progress.stdin.flush()
            
            return True
        except Exception as e:
            self.show_dialog('Ошибка', f'Ошибка сохранения:\n{str(e)}', 'error')
            return False

    def run(self):
        """Основной цикл приложения"""
        # Проверка зависимостей
        try:
            import OpenSSL
            import pandas
        except ImportError as e:
            self.show_dialog(
                'Ошибка', 
                'Необходимо установить:\nsudo apt install python3-pip\npip install pyopenssl pandas openpyxl',
                'error', 500
            )
            return
        
        # Инструкция
        self.show_dialog(
            'Инструкция',
            '1. Выберите папку с сертификатами (.cer)\n'
            '2. Укажите файл для сохранения отчета (.xlsx)\n'
            '3. Дождитесь завершения обработки',
            width=500
        )
        
        # Выбор директории
        directory = self.get_directory()
        if not directory:
            return
        
        # Выбор файла для сохранения
        output_file = self.get_output_file()
        if not output_file:
            return
        
        if not output_file.endswith('.xlsx'):
            output_file += '.xlsx'
        
        # Обработка
        results = self.process_directory(directory)
        
        # Сохранение
        if results:
            if self.save_to_excel(results, output_file):
                self.show_dialog(
                    'Готово',
                    f'Отчет сохранен:\n{output_file}\n\n'
                    f'Обработано сертификатов: {len(results)}',
                    width=500
                )

if __name__ == "__main__":
    app = CertificateParserGUI()
    app.run()