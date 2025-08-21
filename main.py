#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Программа для обработки данных из Excel файлов
Автор: AI Assistant
Версия: 1.0.0
Дата создания: 2024
"""

import os
import sys
import time
import logging
import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path
import traceback

# =============================================================================
# КОНСТАНТЫ И НАСТРОЙКИ ПРОГРАММЫ
# =============================================================================

# Путь к рабочей папке (сырая строка для избежания экранирования)
WORK_DIR = r"/Users/orionflash/Desktop/MyProject/Reaction_Effectiv_LOAD/WORK"

# Названия подпапок
INPUT_FOLDER = "INPUT"      # Папка с входными файлами
OUTPUT_FOLDER = "OUTPUT"    # Папка с выходными файлами
LOGS_FOLDER = "LOGS"        # Папка с логами

# Настройки входных файлов (имя без расширения, расширение отдельно)
INPUT_FILES = [
    {"name": "data1", "extension": ".xlsx"},
    {"name": "data2", "extension": ".xlsx"}
]

# Настройки выходных файлов
OUTPUT_FILES = [
    {"name": "processed_data", "extension": ".csv", "suffix_format": "_YYYYMMDD-HHMMSS"},
    {"name": "processed_data", "extension": ".xlsx", "suffix_format": "_YYYYMMDD-HHMMSS"}
]

# Настройки лог-файла
LOG_FILE = {
    "name": "processing_log",
    "extension": ".log",
    "suffix_format": "_YYYYMMDD"
}

# Режим работы программы
# "process" - обработка данных (основная работа)
# "create-test" - создание тестовых данных
PROGRAM_MODE = "process"

# Уровень логирования (INFO или DEBUG)
LOG_LEVEL = "INFO"

# Сообщения для логирования
LOG_MESSAGES = {
    "start": "Программа запущена",
    "end": "Программа завершена",
    "processing_start": "Начало обработки данных",
    "processing_end": "Обработка данных завершена",
    "file_loaded": "Файл {} загружен успешно",
    "file_saved": "Файл {} сохранен успешно",
    "error": "Ошибка: {}",
    "summary": "Сводка выполнения: {}",
    "time_elapsed": "Время выполнения: {:.2f} секунд",
    "files_processed": "Обработано файлов: {}",
    "outputs_created": "Создано выходных файлов: {}",
    "errors_count": "Количество ошибок: {}",
    "test_data_created": "Тестовые данные созданы успешно",
    "test_file_created": "Создан тестовый файл: {}",
    "test_data_info": "Тестовый файл содержит {} строк и {} столбцов"
}

# =============================================================================
# КЛАСС ДЛЯ ЛОГИРОВАНИЯ
# =============================================================================

class DataProcessorLogger:
    """Класс для управления логированием программы"""
    
    def __init__(self, log_dir, log_name, log_extension, suffix_format, level="INFO"):
        """
        Инициализация логгера
        
        Args:
            log_dir (str): Директория для логов
            log_name (str): Имя лог-файла без расширения
            log_extension (str): Расширение лог-файла
            suffix_format (str): Формат суффикса для имени файла
            level (str): Уровень логирования (INFO/DEBUG)
        """
        self.log_dir = Path(log_dir)
        self.log_name = log_name
        self.log_extension = log_extension
        self.suffix_format = suffix_format
        self.level = level.upper()
        
        # Создаем директорию для логов если её нет
        self.log_dir.mkdir(parents=True, exist_ok=True)
        
        # Формируем имя лог-файла с суффиксом
        timestamp = datetime.now().strftime(suffix_format.replace("YYYY", "%Y").replace("MM", "%m").replace("DD", "%d"))
        self.log_filename = f"{log_name}-{level}{timestamp}{log_extension}"
        self.log_filepath = self.log_dir / self.log_filename
        
        # Настраиваем логирование
        self._setup_logging()
    
    def _setup_logging(self):
        """Настройка системы логирования"""
        # Определяем уровень логирования
        log_level = logging.DEBUG if self.level == "DEBUG" else logging.INFO
        
        # Настраиваем форматтер
        formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # Настраиваем файловый хендлер
        file_handler = logging.FileHandler(self.log_filepath, encoding='utf-8')
        file_handler.setLevel(log_level)
        file_handler.setFormatter(formatter)
        
        # Настраиваем консольный хендлер
        console_handler = logging.StreamHandler()
        console_handler.setLevel(log_level)
        console_handler.setFormatter(formatter)
        
        # Настраиваем корневой логгер
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(log_level)
        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)
        
        # Очищаем хендлеры чтобы избежать дублирования
        self.logger.handlers.clear()
        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)
    
    def log_info(self, message):
        """Логирование информационного сообщения"""
        self.logger.info(message)
    
    def log_debug(self, message):
        """Логирование отладочного сообщения"""
        self.logger.debug(message)
    
    def log_error(self, message):
        """Логирование ошибки"""
        self.logger.error(message)
    
    def log_start(self):
        """Логирование начала работы программы"""
        self.log_info(LOG_MESSAGES["start"])
    
    def log_end(self):
        """Логирование завершения работы программы"""
        self.log_info(LOG_MESSAGES["end"])

# =============================================================================
# КЛАСС ДЛЯ СОЗДАНИЯ ТЕСТОВЫХ ДАННЫХ
# =============================================================================

class TestDataGenerator:
    """Класс для создания тестовых Excel файлов"""
    
    def __init__(self, work_dir, logger):
        """
        Инициализация генератора тестовых данных
        
        Args:
            work_dir (str): Рабочая директория
            logger (DataProcessorLogger): Объект логгера
        """
        self.work_dir = Path(work_dir)
        self.logger = logger
        self.start_time = None
        self.errors_count = 0
        self.files_created = 0
        
        # Создаем необходимые директории
        self._create_directories()
    
    def _create_directories(self):
        """Создание необходимых директорий"""
        directories = [
            self.work_dir / INPUT_FOLDER,
            self.work_dir / OUTPUT_FOLDER,
            self.work_dir / LOGS_FOLDER
        ]
        
        for directory in directories:
            directory.mkdir(parents=True, exist_ok=True)
            self.logger.log_debug(f"Директория {directory} готова к работе")
    
    def create_sample_data(self):
        """Создание тестовых данных для демонстрации"""
        self.start_time = time.time()
        
        try:
            # Создаем первый набор данных
            data1 = {
                'ID': range(1, 101),
                'Имя': [f'Пользователь_{i}' for i in range(1, 101)],
                'Возраст': np.random.randint(18, 65, 100),
                'Город': np.random.choice(['Москва', 'СПб', 'Новосибирск', 'Екатеринбург'], 100),
                'Зарплата': np.random.randint(30000, 150000, 100),
                'Отдел': np.random.choice(['IT', 'Маркетинг', 'Продажи', 'HR'], 100)
            }
            
            # Создаем второй набор данных
            data2 = {
                'ID': range(51, 151),
                'Имя': [f'Сотрудник_{i}' for i in range(51, 151)],
                'Возраст': np.random.randint(20, 70, 100),
                'Город': np.random.choice(['Москва', 'СПб', 'Казань', 'Нижний Новгород'], 100),
                'Зарплата': np.random.randint(25000, 200000, 100),
                'Отдел': np.random.choice(['IT', 'Аналитика', 'Финансы', 'Юридический'], 100)
            }
            
            # Создаем DataFrame'ы
            df1 = pd.DataFrame(data1)
            df2 = pd.DataFrame(data2)
            
            # Сохраняем файлы
            file1_path = self.work_dir / INPUT_FOLDER / "data1.xlsx"
            file2_path = self.work_dir / INPUT_FOLDER / "data2.xlsx"
            
            df1.to_excel(file1_path, index=False, engine='openpyxl')
            df2.to_excel(file2_path, index=False, engine='openpyxl')
            
            self.logger.log_info(LOG_MESSAGES["test_file_created"].format(file1_path.name))
            self.logger.log_debug(LOG_MESSAGES["test_data_info"].format(len(df1), len(df1.columns)))
            self.files_created += 1
            
            self.logger.log_info(LOG_MESSAGES["test_file_created"].format(file2_path.name))
            self.logger.log_debug(LOG_MESSAGES["test_data_info"].format(len(df2), len(df2.columns)))
            self.files_created += 1
            
            self.logger.log_info(LOG_MESSAGES["test_data_created"])
            
        except Exception as e:
            error_msg = f"Ошибка при создании тестовых данных: {str(e)}"
            self.logger.log_error(error_msg)
            self.logger.log_debug(f"Детали ошибки: {traceback.format_exc()}")
            self.errors_count += 1
        
        finally:
            # Генерируем сводку
            self._generate_summary()
    
    def _generate_summary(self):
        """Генерация сводки создания тестовых данных"""
        end_time = time.time()
        execution_time = end_time - self.start_time
        
        summary = {
            'execution_time': execution_time,
            'files_created': self.files_created,
            'errors_count': self.errors_count
        }
        
        # Логируем сводку
        self.logger.log_info(LOG_MESSAGES["summary"].format(summary))
        self.logger.log_info(LOG_MESSAGES["time_elapsed"].format(execution_time))
        self.logger.log_info(f"Создано тестовых файлов: {self.files_created}")
        self.logger.log_info(LOG_MESSAGES["errors_count"].format(self.errors_count))
        
        return summary

# =============================================================================
# КЛАСС ДЛЯ ОБРАБОТКИ ДАННЫХ
# =============================================================================

class DataProcessor:
    """Основной класс для обработки данных"""
    
    def __init__(self, work_dir, logger):
        """
        Инициализация процессора данных
        
        Args:
            work_dir (str): Рабочая директория
            logger (DataProcessorLogger): Объект логгера
        """
        self.work_dir = Path(work_dir)
        self.logger = logger
        self.start_time = None
        self.errors_count = 0
        self.files_processed = 0
        self.outputs_created = 0
        
        # Создаем необходимые директории
        self._create_directories()
    
    def _create_directories(self):
        """Создание необходимых директорий"""
        directories = [
            self.work_dir / INPUT_FOLDER,
            self.work_dir / OUTPUT_FOLDER,
            self.work_dir / LOGS_FOLDER
        ]
        
        for directory in directories:
            directory.mkdir(parents=True, exist_ok=True)
            self.logger.log_debug(f"Директория {directory} готова к работе")
    
    def load_excel_files(self):
        """
        Загрузка данных из Excel файлов
        
        Returns:
            list: Список загруженных DataFrame'ов
        """
        dataframes = []
        
        for file_config in INPUT_FILES:
            try:
                file_path = self.work_dir / INPUT_FOLDER / f"{file_config['name']}{file_config['extension']}"
                
                if file_path.exists():
                    # Загружаем Excel файл
                    df = pd.read_excel(file_path)
                    dataframes.append({
                        'name': file_config['name'],
                        'data': df,
                        'file_path': file_path
                    })
                    
                    self.logger.log_info(LOG_MESSAGES["file_loaded"].format(file_path.name))
                    self.logger.log_debug(f"Загружено {len(df)} строк и {len(df.columns)} столбцов из {file_path.name}")
                    self.files_processed += 1
                else:
                    self.logger.log_error(f"Файл {file_path} не найден")
                    self.errors_count += 1
                    
            except Exception as e:
                error_msg = f"Ошибка при загрузке файла {file_config['name']}: {str(e)}"
                self.logger.log_error(error_msg)
                self.logger.log_debug(f"Детали ошибки: {traceback.format_exc()}")
                self.errors_count += 1
        
        return dataframes
    
    def process_data(self, dataframes):
        """
        Обработка загруженных данных
        
        Args:
            dataframes (list): Список загруженных DataFrame'ов
            
        Returns:
            pd.DataFrame: Обработанные данные
        """
        self.logger.log_info(LOG_MESSAGES["processing_start"])
        
        if not dataframes:
            self.logger.log_error("Нет данных для обработки")
            return pd.DataFrame()
        
        try:
            # Объединяем все DataFrame'ы
            combined_df = pd.concat([df['data'] for df in dataframes], ignore_index=True)
            
            # Удаляем дубликаты
            initial_rows = len(combined_df)
            combined_df = combined_df.drop_duplicates()
            duplicates_removed = initial_rows - len(combined_df)
            
            # Обработка пропущенных значений
            missing_values = combined_df.isnull().sum().sum()
            combined_df = combined_df.fillna("N/A")
            
            # Сортировка по первому столбцу
            if len(combined_df.columns) > 0:
                combined_df = combined_df.sort_values(by=combined_df.columns[0])
            
            self.logger.log_debug(f"Обработано данных: {len(combined_df)} строк")
            self.logger.log_debug(f"Удалено дубликатов: {duplicates_removed}")
            self.logger.log_debug(f"Заполнено пропущенных значений: {missing_values}")
            
            self.logger.log_info(LOG_MESSAGES["processing_end"])
            return combined_df
            
        except Exception as e:
            error_msg = f"Ошибка при обработке данных: {str(e)}"
            self.logger.log_error(error_msg)
            self.logger.log_debug(f"Детали ошибки: {traceback.format_exc()}")
            self.errors_count += 1
            return pd.DataFrame()
    
    def save_outputs(self, processed_data):
        """
        Сохранение обработанных данных в выходные файлы
        
        Args:
            processed_data (pd.DataFrame): Обработанные данные
        """
        if processed_data.empty:
            self.logger.log_error("Нет данных для сохранения")
            return
        
        for output_config in OUTPUT_FILES:
            try:
                # Формируем имя файла с суффиксом
                timestamp = datetime.now().strftime(
                    output_config["suffix_format"]
                    .replace("YYYY", "%Y")
                    .replace("MM", "%m")
                    .replace("DD", "%d")
                    .replace("HH", "%H")
                    .replace("MM", "%M")
                    .replace("SS", "%S")
                )
                
                filename = f"{output_config['name']}{timestamp}{output_config['extension']}"
                file_path = self.work_dir / OUTPUT_FOLDER / filename
                
                # Сохраняем файл в зависимости от формата
                if output_config['extension'].lower() == '.csv':
                    # Сохраняем CSV с разделителем ";"
                    processed_data.to_csv(file_path, sep=';', index=False, encoding='utf-8')
                elif output_config['extension'].lower() == '.xlsx':
                    # Сохраняем Excel
                    processed_data.to_excel(file_path, index=False, engine='openpyxl')
                
                self.logger.log_info(LOG_MESSAGES["file_saved"].format(filename))
                self.logger.log_debug(f"Файл сохранен: {file_path}")
                self.outputs_created += 1
                
            except Exception as e:
                error_msg = f"Ошибка при сохранении файла {output_config['name']}: {str(e)}"
                self.logger.log_error(error_msg)
                self.logger.log_debug(f"Детали ошибки: {traceback.format_exc()}")
                self.errors_count += 1
    
    def generate_summary(self):
        """Генерация сводки выполнения программы"""
        end_time = time.time()
        execution_time = end_time - self.start_time
        
        summary = {
            'execution_time': execution_time,
            'files_processed': self.files_processed,
            'outputs_created': self.outputs_created,
            'errors_count': self.errors_count
        }
        
        # Логируем сводку
        self.logger.log_info(LOG_MESSAGES["summary"].format(summary))
        self.logger.log_info(LOG_MESSAGES["time_elapsed"].format(execution_time))
        self.logger.log_info(LOG_MESSAGES["files_processed"].format(self.files_processed))
        self.logger.log_info(LOG_MESSAGES["outputs_created"].format(self.outputs_created))
        self.logger.log_info(LOG_MESSAGES["errors_count"].format(self.errors_count))
        
        return summary
    
    def run(self):
        """Основной метод запуска обработки данных"""
        self.start_time = time.time()
        
        try:
            # Загружаем данные
            dataframes = self.load_excel_files()
            
            # Обрабатываем данные
            processed_data = self.process_data(dataframes)
            
            # Сохраняем результаты
            self.save_outputs(processed_data)
            
        except Exception as e:
            error_msg = f"Критическая ошибка в процессе выполнения: {str(e)}"
            self.logger.log_error(error_msg)
            self.logger.log_debug(f"Детали ошибки: {traceback.format_exc()}")
            self.errors_count += 1
        
        finally:
            # Генерируем сводку
            self.generate_summary()

# =============================================================================
# ГЛАВНАЯ ФУНКЦИЯ
# =============================================================================

def main():
    """Главная функция программы"""
    
    try:
        # Создаем логгер
        logger = DataProcessorLogger(
            log_dir=Path(WORK_DIR) / LOGS_FOLDER,
            log_name=LOG_FILE["name"],
            log_extension=LOG_FILE["extension"],
            suffix_format=LOG_FILE["suffix_format"],
            level=LOG_LEVEL
        )
        
        # Логируем начало работы
        logger.log_start()
        
        if PROGRAM_MODE == 'create-test':
            # Режим создания тестовых данных
            logger.log_info("Режим: Создание тестовых данных")
            generator = TestDataGenerator(WORK_DIR, logger)
            generator.create_sample_data()
            print("Тестовые данные созданы успешно. Проверьте папку INPUT.")
            
        else:
            # Режим обработки данных (по умолчанию)
            logger.log_info("Режим: Обработка данных")
            processor = DataProcessor(WORK_DIR, logger)
            processor.run()
            print("Программа выполнена успешно. Проверьте папку OUTPUT для результатов.")
        
        # Логируем завершение работы
        logger.log_end()
        
    except Exception as e:
        print(f"Критическая ошибка при запуске программы: {str(e)}")
        sys.exit(1)

# =============================================================================
# ТОЧКА ВХОДА
# =============================================================================

if __name__ == "__main__":
    main()
