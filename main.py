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
    {"name": "data1_20250821_144017", "extension": ".xlsx"},
    {"name": "data2_20250821_144017", "extension": ".xlsx"}
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
LOG_LEVEL = "DEBUG"

# Параметры генерации тестовых данных
DATA_PARAMS = {
    "total_employees": 1600,        # Общее количество сотрудников
    "effective_share": 0.80,        # Доля эффективных сотрудников (80%)
    "operational_income_july_min": 500000,    # Минимальный операционный доход на 31 июля 2025 (тыс. руб.)
    "operational_income_july_max": 20000000,   # Максимальный операционный доход на 31 июля 2025 (тыс. руб.)
    "operational_income_august_min": 500000,  # Минимальный операционный доход на 20 августа 2025 (тыс. руб.)
    "operational_income_august_max": 220000000, # Максимальный операционный доход на 20 августа 2025 (тыс. руб.)
    "employee_overlap": 0.90,       # Доля одинаковых сотрудников в двух файлах (90%)
    "new_employees_share": 0.05,    # Доля новых сотрудников (5%)
    "removed_employees_share": 0.05 # Доля убранных сотрудников (5%)
}

# Территориальные банки (ТБ)
TERRITORIAL_BANKS = [
    "Байкальский банк",
    "Волго-Вятский банк",
    "Дальневосточный банк",
    "Московский банк",
    "Поволжский банк",
    "Северо-Западный банк",
    "Сибирский банк",
    "Среднерусский банк",
    "Уральский банк",
    "Центрально-Черноземный банк",
    "Юго-Западный банк"
]

# Головные отделения (ГОСБ) - 106 названий
HEAD_OFFICES = [
    # Байкальский банк (5 ГОСБ)
    "Аппарат Байкальского Банка",
    "Бурятское ГОСБ №8601",
    "Иркутское ГОСБ №8586",
    "Читинское ГОСБ №8600",
    "Якутское ГОСБ №8603",
    
    # Волго-Вятский банк (12 ГОСБ)
    "Аппарат Волго-Вятского банка",
    "Банк Татарстан ГОСБ №8610",
    "Владимирское ГОСБ №8611",
    "Головное отделение по Нижегородской области",
    "Кировское ГОСБ №8612",
    "Марий Эл ГОСБ №8614",
    "Мордовское ГОСБ №8589",
    "Пермское ГОСБ №6984",
    "Удмуртское ГОСБ №8618",
    "Чувашское ГОСБ №8613",
    
    # Дальневосточный банк (9 ГОСБ)
    "Аппарат Дальневосточного банка",
    "Биробиджанское ГОСБ №4157",
    "Благовещенское ГОСБ №8636",
    "Камчатское ГОСБ №8556",
    "Приморское ГОСБ №8635",
    "Северо-Восточное ГОСБ №8645",
    "Хабаровское ГОСБ №9070",
    "Чукотское ГОСБ №8557",
    "Южно-Сахалинское ГОСБ №8567",
    
    # Московский банк (5 ГОСБ)
    "УПРАВЛЕНИЕ по работе с предприятиями инфраструктуры",
    "УПРАВЛЕНИЕ по работе с предприятиями промышленности",
    "УПРАВЛЕНИЕ по работе с предприятиями сферы недвижимости",
    "УПРАВЛЕНИЕ по работе с предприятиями сферы услуг",
    "УПРАВЛЕНИЕ по работе с предприятиями торговли",
    
    # Поволжский банк (8 ГОСБ)
    "Аппарат Поволжского банка",
    "Астраханское ГОСБ №8625",
    "Волгоградское ГОСБ №8621",
    "Оренбургское ГОСБ №8623",
    "Пензенское ГОСБ №8624",
    "Самарское ГОСБ №6991",
    "Саратовское ГОСБ №8622",
    "Ульяновское ГОСБ №8588",
    
    # Северо-Западный банк (10 ГОСБ)
    "Аппарат Северо-Западного банка",
    "Архангельское ГОСБ №8637",
    "Вологодское ГОСБ №8638",
    "ГО по г. Санкт-Петербургу",
    "ГО по Ленинградской области",
    "Калининградское ГОСБ №8626",
    "Карельское ГОСБ №8628",
    "Коми ГОСБ №8617",
    "Мурманское ГОСБ №8627",
    "Новгородское ГОСБ №8629",
    "Псковское ГОСБ №8630",
    
    # Сибирский банк (8 ГОСБ)
    "Абаканское ГОСБ №8602",
    "Алтайское ГОСБ №8644",
    "Аппарат Сибирского банка",
    "Кемеровское ГОСБ №8615",
    "Красноярское ГОСБ №8646",
    "Новосибирское ГОСБ №8047",
    "Омское ГОСБ №8634",
    "Томское ГОСБ №8616",
    
    # Среднерусский банк (13 ГОСБ)
    "Аппарат Среднерусского банка",
    "Брянское ГОСБ №8605",
    "Восточное ГОСБ №1023",
    "Западное ГОСБ №1025",
    "Ивановское ГОСБ №8639",
    "Калужское ГОСБ №8608",
    "Костромское ГОСБ №8640",
    "Рязанское ГОСБ №8606",
    "Северное ГОСБ №1026",
    "Смоленское ГОСБ №8609",
    "Тверское ГОСБ №8607",
    "Тульское ГОСБ №8604",
    "Южное ГОСБ №1024",
    "Ярославское ГОСБ №17",
    
    # Уральский банк (7 ГОСБ)
    "Аппарат Уральского банка",
    "Башкирское ГОСБ №8598",
    "Западно-Сибирское ГОСБ №8647",
    "Курганское ГОСБ №8599",
    "Свердловское ГОСБ №7003",
    "Челябинское ГОСБ №8597",
    "Югорское отделение №5940",
    "Ямало-Ненецкое отделение №8369",
    
    # Центрально-Черноземный банк (8 ГОСБ)
    "Аппарат Центрально-Черноземного банка",
    "Белгородское ГОСБ №8592",
    "ГО по Воронежской области",
    "Головное отделение по ЛНР",
    "Курское ГОСБ №8596",
    "Липецкое ГОСБ №8593",
    "Орловское ГОСБ №8595",
    "Тамбовское ГОСБ №8594",
    
    # Юго-Западный банк (12 ГОСБ)
    "Адыгейское ОСБ №8620",
    "Аппарат Юго-Западного банка",
    "Головное отделение по ДНР",
    "Головное отделение по Республике Крым",
    "Ингушское отделение № 8633",
    "Кабардино-Балкарское ОСБ №8631",
    "Калмыцкое ОСБ №8579",
    "Карачаево-Черкесское ОСБ №8585",
    "Краснодарское ГОСБ №8619",
    "Ростовское ГОСБ №5221",
    "Северо-Осетинское ОСБ №8632",
    "Ставропольское ГОСБ №5230",
    "Чеченское ОСБ №8643"
]

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
    "test_data_info": "Тестовый файл содержит {} строк и {} столбцов",
    "data_generation_start": "Начало генерации тестовых данных",
    "data_generation_end": "Генерация тестовых данных завершена",
    "employees_created": "Создано сотрудников: {}",
    "tb_distribution": "Распределение по ТБ: {}",
    "gosb_distribution": "Распределение по ГОСБ: {}",
    "effective_distribution": "Распределение эффективных: {}",
    "directory_ready": "Директория {} готова к работе",
    "tb_mapping_created": "Создано распределение ГОСБ по ТБ: {} ТБ",
    "progress_employees": "Сгенерировано сотрудников: {}",
    "unique_tn_fio": "Уникальных ТН: {}, Уникальных ФИО: {}",
    "duplicate_tn_error": "ОШИБКА: Дублирование табельных номеров!",
    "duplicate_fio_error": "ОШИБКА: Дублирование ФИО!",
    "file_saved_debug": "Файл сохранен: {}",
    "file_size_debug": "Размер файла: {} строк, {} столбцов",
    "generation_error": "Ошибка при генерации тестовых данных: {}",
    "save_error": "Ошибка при сохранении файла: {}",
    "details_error": "Детали ошибки: {}",
    "file_not_found": "Файл {} не найден",
    "load_file_error": "Ошибка при загрузке файла {}: {}",
    "no_data_to_process": "Нет данных для обработки",
    "no_data_to_save": "Нет данных для сохранения",
    "data_processed": "Обработано данных: {} строк",
    "rows_columns_loaded": "Загружено {} строк и {} столбцов из файла {}",
    "duplicates_removed": "Удалено дубликатов: {}",
    "missing_values_filled": "Заполнено пропущенных значений: {}",
    "processing_error": "Ошибка при обработке данных: {}",
    "critical_error": "Критическая ошибка в процессе выполнения: {}",
    "mode_create_test": "Режим: Создание тестовых данных",
    "mode_process": "Режим: Обработка данных",
    "test_data_success": "Тестовые данные созданы успешно. Проверьте папку INPUT.",
    "process_success": "Программа выполнена успешно. Проверьте папку OUTPUT для результатов.",
    "main_critical_error": "Критическая ошибка при запуске программы: {}",
    "file_saved_debug_old": "Файл сохранен: {}",
    "analysis_file1": "=== Анализ файла 1 (31 июля 2025 года) ===",
    "analysis_file2": "=== Анализ файла 2 (20 августа 2025 года) ===",
    "analysis_overlap": "=== Анализ перекрытия сотрудников ===",
    "employees_2024": "Сотрудников в 2024: {}",
    "employees_2025": "Сотрудников в 2025: {}",
    "overlap_info": "Перекрытие (одинаковые): {} ({:.1f}%)",
    "new_employees_info": "Новых сотрудников: {} ({:.1f}%)",
    "removed_employees_info": "Убранных сотрудников: {} ({:.1f}%)",
    "unique_tn_fio_file1": "Файл 1 - Уникальных ТН: {}, Уникальных ФИО: {}",
    "unique_tn_fio_file2": "Файл 2 - Уникальных ТН: {}, Уникальных ФИО: {}"
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
    """Класс для создания тестовых данных для анализа эффективности"""
    
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
        self.employees_created = 0
        
        # Создаем необходимые директории
        self._create_directories()
        
        # Создаем распределение ГОСБ по ТБ
        self._create_tb_gosb_mapping()
    
    def _create_directories(self):
        """Создание необходимых директорий"""
        directories = [
            self.work_dir / INPUT_FOLDER,
            self.work_dir / OUTPUT_FOLDER,
            self.work_dir / LOGS_FOLDER
        ]
        
        for directory in directories:
            directory.mkdir(parents=True, exist_ok=True)
            self.logger.log_debug(LOG_MESSAGES["directory_ready"].format(directory))
    
    def _create_tb_gosb_mapping(self):
        """Создание распределения ГОСБ по ТБ"""
        self.tb_gosb_mapping = {}
        current_index = 0
        
        for tb in TERRITORIAL_BANKS:
            # Определяем количество ГОСБ для данного ТБ
            if tb == "Байкальский банк":
                gosb_count = 5
            elif tb == "Волго-Вятский банк":
                gosb_count = 12
            elif tb == "Дальневосточный банк":
                gosb_count = 9
            elif tb == "Московский банк":
                gosb_count = 5
            elif tb == "Поволжский банк":
                gosb_count = 8
            elif tb == "Северо-Западный банк":
                gosb_count = 10
            elif tb == "Сибирский банк":
                gosb_count = 8
            elif tb == "Среднерусский банк":
                gosb_count = 13
            elif tb == "Уральский банк":
                gosb_count = 7
            elif tb == "Центрально-Черноземный банк":
                gosb_count = 8
            elif tb == "Юго-Западный банк":
                gosb_count = 12
            else:
                gosb_count = 10
            
            # Берем ГОСБ для данного ТБ
            self.tb_gosb_mapping[tb] = HEAD_OFFICES[current_index:current_index + gosb_count]
            current_index += gosb_count
        
        self.logger.log_debug(LOG_MESSAGES["tb_mapping_created"].format(len(self.tb_gosb_mapping)))
    
    def _generate_tn(self):
        """Генерация табельного номера (10 знаков, справа значащие)"""
        # Генерируем случайное число от 1 до 9999999999
        tn_number = np.random.randint(1, 10000000000)
        # Форматируем в 10 знаков с ведущими нулями
        return f"{tn_number:010d}"
    
    def _generate_fio(self):
        """Генерация уникального ФИО"""
        # Списки имен, фамилий и отчеств
        first_names_male = [
            "Александр", "Сергей", "Владимир", "Дмитрий", "Андрей", "Алексей", "Максим", "Иван", "Михаил", "Николай",
            "Артем", "Денис", "Евгений", "Даниил", "Роман", "Тимур", "Владислав", "Павел", "Константин", "Игорь"
        ]
        
        first_names_female = [
            "Анна", "Мария", "Елена", "Ольга", "Татьяна", "Наталья", "Ирина", "Светлана", "Юлия", "Екатерина",
            "Анастасия", "Дарья", "Ксения", "Виктория", "Полина", "Алиса", "София", "Вероника", "Арина", "Диана"
        ]
        
        last_names = [
            "Иванов", "Смирнов", "Кузнецов", "Попов", "Васильев", "Петров", "Соколов", "Михайлов", "Новиков", "Федоров",
            "Морозов", "Волков", "Алексеев", "Лебедев", "Семенов", "Егоров", "Павлов", "Козлов", "Степанов", "Николаев"
        ]
        
        middle_names_male = [
            "Александрович", "Сергеевич", "Владимирович", "Дмитриевич", "Андреевич", "Алексеевич", "Максимович", "Иванович", "Михайлович", "Николаевич"
        ]
        
        middle_names_female = [
            "Александровна", "Сергеевна", "Владимировна", "Дмитриевна", "Андреевна", "Алексеевна", "Максимовна", "Ивановна", "Михайловна", "Николаевна"
        ]
        
        # Выбираем пол случайно
        is_male = np.random.choice([True, False])
        
        if is_male:
            first_name = np.random.choice(first_names_male)
            middle_name = np.random.choice(middle_names_male)
        else:
            first_name = np.random.choice(first_names_female)
            middle_name = np.random.choice(middle_names_female)
        
        last_name = np.random.choice(last_names)
        
        return f"{last_name} {first_name} {middle_name}"
    
    def _generate_effective_status(self):
        """Генерация статуса эффективности"""
        # 80% эффективных, 20% неэффективных
        is_effective = np.random.random() < DATA_PARAMS["effective_share"]
        return "👍" if is_effective else "👎"
    
    def _generate_operational_income_data(self):
        """Генерация данных об операционном доходе"""
        # Операционный доход на 31 июля 2025
        income_july = np.random.randint(DATA_PARAMS["operational_income_july_min"], DATA_PARAMS["operational_income_july_max"] + 1)
        
        # Операционный доход на 20 августа 2025 (обязательно >= июльского)
        # Минимальный доход = доход на 31 июля, максимальный = заданный максимум
        income_august = np.random.randint(income_july, DATA_PARAMS["operational_income_august_max"] + 1)
        
        # Прирост в процентах
        growth_percent = ((income_august - income_july) / income_july * 100) if income_july > 0 else 0
        
        # Прирост в тыс. руб.
        growth_amount = income_august - income_july
        
        # ОД конец квартала = август 2025
        od_quarter = income_august
        
        return {
            'operational_income_july': income_july,
            'operational_income_august': income_august,
            'growth_percent': round(growth_percent, 2),
            'growth_amount': growth_amount,
            'od_quarter': income_august
        }
    
    def create_sample_data(self):
        """Создание тестовых данных для анализа эффективности"""
        self.start_time = time.time()
        
        try:
            self.logger.log_info(LOG_MESSAGES["data_generation_start"])
            
            # Создаем списки для данных
            data1 = []  # Данные на 31 июля 2025 года
            data2 = []  # Данные на 20 августа 2025 года
            used_fios = set()
            
            # Генерируем базовый список сотрудников для 31 июля 2025 года
            base_employees = []
            for i in range(DATA_PARAMS["total_employees"]):
                # Генерируем уникальный ФИО
                while True:
                    fio = self._generate_fio()
                    if fio not in used_fios:
                        used_fios.add(fio)
                        break
                
                # Выбираем случайный ТБ и ГОСБ
                tb = np.random.choice(TERRITORIAL_BANKS)
                gosb = np.random.choice(self.tb_gosb_mapping[tb])
                
                # Генерируем остальные данные для 31 июля 2025 года
                tn = self._generate_tn()
                effective_status = self._generate_effective_status()
                
                # Генерируем базовый доход (как бы на начало периода) и доход на 31 июля
                # Базовый доход в диапазоне от 60% до 90% от минимального дохода июля
                base_income_min = int(DATA_PARAMS["operational_income_july_min"] * 0.6)
                base_income_max = int(DATA_PARAMS["operational_income_july_min"] * 0.9)
                base_income = np.random.randint(base_income_min, base_income_max + 1)
                income_july = np.random.randint(DATA_PARAMS["operational_income_july_min"], DATA_PARAMS["operational_income_july_max"] + 1)
                
                # Вычисляем прирост от базового дохода до 31 июля
                growth_percent_july = ((income_july - base_income) / base_income * 100) if base_income > 0 else 0
                growth_amount_july = income_july - base_income
                
                # Создаем строку данных для 31 июля 2025 года
                row_july = {
                    'ИНД (ТБ_ГОСБ_ТН)': f"{tb}_{gosb}_{tn}",
                    'ИНД (ТБ_ГОСБ_ФИО)': f"{tb}_{gosb}_{fio}",
                    'ТН 10': tn,
                    'ТБ': tb,
                    'ГОСБ': gosb,
                    'КМ': fio,
                    'Эффективный КМ': effective_status,
                    '2025, тыс. руб.': income_july,
                    '2024, тыс. руб. на конец месяца': base_income,
                    'Прирост, %': round(growth_percent_july, 2),
                    'Прирост, тыс. руб.': growth_amount_july,
                    'ОД конец квартала, тыс. руб.': income_july
                }
                
                data1.append(row_july)
                base_employees.append({
                    'fio': fio,
                    'tn': tn,
                    'tb': tb,
                    'gosb': gosb,
                    'effective_status': effective_status,
                    'income_july': income_july
                })
                
                # Логируем прогресс каждые 100 сотрудников
                if (i + 1) % 100 == 0:
                    self.logger.log_debug(LOG_MESSAGES["progress_employees"].format(i + 1))
            
            # Теперь создаем данные для 20 августа 2025 года
            # 90% сотрудников остаются, 5% новых, 5% убираем
            overlap_count = int(DATA_PARAMS["total_employees"] * DATA_PARAMS["employee_overlap"])
            new_count = int(DATA_PARAMS["total_employees"] * DATA_PARAMS["new_employees_share"])
            removed_count = DATA_PARAMS["total_employees"] - overlap_count - new_count
            
            # Сотрудники, которые остаются (90%)
            remaining_employees = np.random.choice(base_employees, overlap_count, replace=False)
            
            # Создаем данные для оставшихся сотрудников
            for emp in remaining_employees:
                # Генерируем доход на 20 августа (>= дохода на 31 июля)
                # Минимальный доход = доход на 31 июля, максимальный = заданный максимум
                income_august = np.random.randint(emp['income_july'], DATA_PARAMS["operational_income_august_max"] + 1)
                
                # Вычисляем прирост
                growth_percent = ((income_august - emp['income_july']) / emp['income_july'] * 100) if emp['income_july'] > 0 else 0
                growth_amount = income_august - emp['income_july']
                
                row_august = {
                    'ИНД (ТБ_ГОСБ_ТН)': f"{emp['tb']}_{emp['gosb']}_{emp['tn']}",
                    'ИНД (ТБ_ГОСБ_ФИО)': f"{emp['tb']}_{emp['gosb']}_{emp['fio']}",
                    'ТН 10': emp['tn'],
                    'ТБ': emp['tb'],
                    'ГОСБ': emp['gosb'],
                    'КМ': emp['fio'],
                    'Эффективный КМ': emp['effective_status'],
                    '2025, тыс. руб.': income_august,
                    '2024, тыс. руб. на конец месяца': emp['income_july'],
                    'Прирост, %': round(growth_percent, 2),
                    'Прирост, тыс. руб.': growth_amount,
                    'ОД конец квартала, тыс. руб.': income_august
                }
                
                data2.append(row_august)
            
            # Добавляем новых сотрудников (5%)
            for i in range(new_count):
                # Генерируем уникальный ФИО
                while True:
                    fio = self._generate_fio()
                    if fio not in used_fios:
                        used_fios.add(fio)
                        break
                
                # Выбираем случайный ТБ и ГОСБ
                tb = np.random.choice(TERRITORIAL_BANKS)
                gosb = np.random.choice(self.tb_gosb_mapping[tb])
                
                # Генерируем данные для нового сотрудника
                tn = self._generate_tn()
                effective_status = self._generate_effective_status()
                income_data = self._generate_operational_income_data()
                
                # Создаем строку данных для нового сотрудника
                row_august = {
                    'ИНД (ТБ_ГОСБ_ТН)': f"{tb}_{gosb}_{tn}",
                    'ИНД (ТБ_ГОСБ_ФИО)': f"{tb}_{gosb}_{fio}",
                    'ТН 10': tn,
                    'ТБ': tb,
                    'ГОСБ': gosb,
                    'КМ': fio,
                    'Эффективный КМ': effective_status,
                    '2025, тыс. руб.': income_data['operational_income_august'],
                    '2024, тыс. руб. на конец месяца': income_data['operational_income_july'],
                    'Прирост, %': income_data['growth_percent'],
                    'Прирост, тыс. руб.': income_data['growth_amount'],
                    'ОД конец квартала, тыс. руб.': income_data['od_quarter']
                }
                
                data2.append(row_august)
            
            # Создаем DataFrame'ы
            df1 = pd.DataFrame(data1)  # Данные на 31 июля 2025 года
            df2 = pd.DataFrame(data2)  # Данные на 20 августа 2025 года
            self.employees_created = len(df1) + len(df2)
            
            # Анализируем распределение
            self._analyze_distribution(df1, df2)
            
            # Сохраняем файлы
            self._save_data_files(df1, df2)
            
            self.logger.log_info(LOG_MESSAGES["data_generation_end"])
            
        except Exception as e:
            error_msg = LOG_MESSAGES["generation_error"].format(str(e))
            self.logger.log_error(error_msg)
            self.logger.log_debug(LOG_MESSAGES["details_error"].format(traceback.format_exc()))
            self.errors_count += 1
        
        finally:
            # Генерируем сводку
            self._generate_summary()
    
    def _analyze_distribution(self, df1, df2):
        """Анализ распределения данных по двум файлам"""
        # Анализ файла 1 (31 июля 2025 года)
        self.logger.log_info(LOG_MESSAGES["analysis_file1"])
        tb_distribution_1 = df1['ТБ'].value_counts()
        self.logger.log_info(LOG_MESSAGES["tb_distribution"].format(dict(tb_distribution_1)))
        
        effective_distribution_1 = df1['Эффективный КМ'].value_counts()
        self.logger.log_info(LOG_MESSAGES["effective_distribution"].format(dict(effective_distribution_1)))
        
        # Анализ файла 2 (20 августа 2025 года)
        self.logger.log_info(LOG_MESSAGES["analysis_file2"])
        tb_distribution_2 = df2['ТБ'].value_counts()
        self.logger.log_info(LOG_MESSAGES["tb_distribution"].format(dict(tb_distribution_2)))
        
        effective_distribution_2 = df2['Эффективный КМ'].value_counts()
        self.logger.log_info(LOG_MESSAGES["effective_distribution"].format(dict(effective_distribution_2)))
        
        # Анализ перекрытия сотрудников
        self.logger.log_info(LOG_MESSAGES["analysis_overlap"])
        employees_july = set(df1['КМ'])
        employees_august = set(df2['КМ'])
        
        overlap_employees = employees_july.intersection(employees_august)
        new_employees = employees_august - employees_july
        removed_employees = employees_july - employees_august
        
        self.logger.log_info(f"Сотрудников на 31 июля: {len(employees_july)}")
        self.logger.log_info(f"Сотрудников на 20 августа: {len(employees_august)}")
        self.logger.log_info(f"Перекрытие (одинаковые): {len(overlap_employees)} ({len(overlap_employees)/len(employees_july)*100:.1f}%)")
        self.logger.log_info(f"Новых сотрудников: {len(new_employees)} ({len(new_employees)/len(employees_august)*100:.1f}%)")
        self.logger.log_info(f"Убранных сотрудников: {len(removed_employees)} ({len(removed_employees)/len(employees_july)*100:.1f}%)")
        
        # Проверяем уникальность ТН и ФИО в каждом файле
        unique_tn_1 = df1['ТН 10'].nunique()
        unique_fio_1 = df1['КМ'].nunique()
        unique_tn_2 = df2['ТН 10'].nunique()
        unique_fio_2 = df2['КМ'].nunique()
        
        self.logger.log_debug(f"Файл 1 (31 июля) - Уникальных ТН: {unique_tn_1}, Уникальных ФИО: {unique_fio_1}")
        self.logger.log_debug(f"Файл 2 (20 августа) - Уникальных ТН: {unique_tn_2}, Уникальных ФИО: {unique_fio_2}")
        
        if unique_tn_1 != len(df1):
            self.logger.log_error(LOG_MESSAGES["duplicate_tn_error"])
        if unique_fio_1 != len(df1):
            self.logger.log_error(LOG_MESSAGES["duplicate_fio_error"])
        if unique_tn_2 != len(df2):
            self.logger.log_error(LOG_MESSAGES["duplicate_tn_error"])
        if unique_fio_2 != len(df2):
            self.logger.log_error(LOG_MESSAGES["duplicate_tn_error"])
    
    def _save_data_files(self, df1, df2):
        """Сохранение данных в два файла"""
        try:
            # Формируем имена файлов с временной меткой
            timestamp = datetime.now().strftime("_%Y%m%d_%H%M%S")
            filename1 = f"data1{timestamp}.xlsx"  # Файл на 31 июля 2025 года
            filename2 = f"data2{timestamp}.xlsx"  # Файл на 20 августа 2025 года
            
            file_path1 = self.work_dir / INPUT_FOLDER / filename1
            file_path2 = self.work_dir / INPUT_FOLDER / filename2
            
            # Сохраняем файл 1 (31 июля 2025 года)
            df1.to_excel(file_path1, index=False, engine='openpyxl')
            self.logger.log_info(LOG_MESSAGES["test_file_created"].format(filename1))
            self.logger.log_debug(LOG_MESSAGES["file_saved_debug"].format(file_path1))
            self.logger.log_debug(LOG_MESSAGES["file_size_debug"].format(len(df1), len(df1.columns)))
            self.files_created += 1
            
            # Сохраняем файл 2 (20 августа 2025 года)
            df2.to_excel(file_path2, index=False, engine='openpyxl')
            self.logger.log_info(LOG_MESSAGES["test_file_created"].format(filename2))
            self.logger.log_debug(LOG_MESSAGES["file_saved_debug"].format(file_path2))
            self.logger.log_debug(LOG_MESSAGES["file_size_debug"].format(len(df2), len(df2.columns)))
            self.files_created += 1
            
        except Exception as e:
            error_msg = LOG_MESSAGES["save_error"].format(str(e))
            self.logger.log_error(error_msg)
            self.logger.log_debug(LOG_MESSAGES["details_error"].format(traceback.format_exc()))
            self.errors_count += 1
    
    def _generate_summary(self):
        """Генерация сводки создания тестовых данных"""
        end_time = time.time()
        execution_time = end_time - self.start_time
        
        summary = {
            'execution_time': execution_time,
            'employees_created': self.employees_created,
            'files_created': self.files_created,
            'errors_count': self.errors_count
        }
        
        # Логируем сводку
        self.logger.log_info(LOG_MESSAGES["summary"].format(summary))
        self.logger.log_info(LOG_MESSAGES["time_elapsed"].format(execution_time))
        self.logger.log_info(LOG_MESSAGES["employees_created"].format(self.employees_created))
        self.logger.log_info(LOG_MESSAGES["outputs_created"].format(self.files_created))
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
            self.logger.log_debug(LOG_MESSAGES["directory_ready"].format(directory))
    
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
                    self.logger.log_debug(LOG_MESSAGES["rows_columns_loaded"].format(len(df), len(df.columns), file_path.name))
                    self.files_processed += 1
                else:
                    self.logger.log_error(LOG_MESSAGES["file_not_found"].format(file_path))
                    self.errors_count += 1
                    
            except Exception as e:
                error_msg = LOG_MESSAGES["load_file_error"].format(file_config['name'], str(e))
                self.logger.log_error(error_msg)
                self.logger.log_debug(LOG_MESSAGES["details_error"].format(traceback.format_exc()))
                self.errors_count += 1
        
        return dataframes
    
    def process_data(self, dataframes):
        """
        Обработка загруженных данных с объединением и расчетом новых колонок
        
        Args:
            dataframes (list): Список загруженных DataFrame'ов
            
        Returns:
            pd.DataFrame: Обработанные данные
        """
        self.logger.log_info(LOG_MESSAGES["processing_start"])
        
        if not dataframes:
            self.logger.log_error(LOG_MESSAGES["no_data_to_process"])
            return pd.DataFrame()
        
        try:
            # Находим файлы data1 и data2 по именам из INPUT_FILES
            df1 = None
            df2 = None
            
            # Получаем имена файлов из конфигурации (без расширения)
            file1_name = INPUT_FILES[0]['name']
            file2_name = INPUT_FILES[1]['name']
            
            for df_info in dataframes:
                if df_info['name'] == file1_name:
                    df1 = df_info['data']
                elif df_info['name'] == file2_name:
                    df2 = df_info['data']
            
            if df1 is None or df2 is None:
                self.logger.log_error(f"Не найдены файлы {file1_name}.xlsx или {file2_name}.xlsx")
                return pd.DataFrame()
            
            self.logger.log_debug(f"Загружены файлы: data1 ({len(df1)} строк), data2 ({len(df2)} строк)")
            
            # Создаем список уникальных значений ТН 10, ТБ, ГОСБ, ФИО
            # Объединяем все уникальные ТН из обоих файлов
            all_tn = pd.concat([
                df1[['ТН 10', 'ТБ', 'ГОСБ', 'КМ']].drop_duplicates(),
                df2[['ТН 10', 'ТБ', 'ГОСБ', 'КМ']].drop_duplicates()
            ]).drop_duplicates(subset=['ТН 10'], keep='last')
            
            self.logger.log_debug(f"Создан список из {len(all_tn)} уникальных ТН")
            
            # Создаем результирующий DataFrame
            result_data = []
            
            for _, row in all_tn.iterrows():
                tn = row['ТН 10']
                tb = row['ТБ']
                gosb = row['ГОСБ']
                fio = row['КМ']
                
                # Ищем данные в файле 1
                data1_row = df1[df1['ТН 10'] == tn]
                data2_row = df2[df2['ТН 10'] == tn]
                
                # Получаем значения из файла 1
                od_current = data1_row['2025, тыс. руб.'].iloc[0] if len(data1_row) > 0 else 0
                od_previous = data1_row['2024, тыс. руб. на конец месяца'].iloc[0] if len(data1_row) > 0 else 0
                
                # Получаем эффективность из файла 2, если нет - из файла 1
                if len(data2_row) > 0:
                    effectiveness = data2_row['Эффективный КМ'].iloc[0]
                elif len(data1_row) > 0:
                    effectiveness = data1_row['Эффективный КМ'].iloc[0]
                else:
                    effectiveness = "👎"
                
                # Конвертируем эффективность в числовое значение
                effectiveness_num = 1 if effectiveness == "👍" else 0
                
                # Рассчитываем темп ОД
                if od_previous == 0:
                    if od_current > 0:
                        temp_od = 100
                    elif od_current < 0:
                        temp_od = -100
                    else:
                        temp_od = 0
                else:
                    temp_od = (od_current - od_previous) / abs(od_previous) * 100
                
                # Создаем строку результата
                result_row = {
                    'ТН 10': tn,
                    'ТБ': tb,
                    'ГОСБ': gosb,
                    'ФИО': fio,
                    'ИНД (ТБ_ГОСБ_ТН)': f"{tb}_{gosb}_{tn}",
                    'ИНД (ТБ_ГОСБ_ФИО)': f"{tb}_{gosb}_{fio}",
                    'ИНД (ТБ_ГОСБ)': f"{tb}_{gosb}",
                    'ОД ТЕКУЩИЙ': od_current,
                    'ОД ПРОШЛЫЙ': od_previous,
                    'ЭФФЕКТИВНОСТЬ': effectiveness_num,
                    'ТЕМП ОД': round(temp_od, 2)
                }
                
                result_data.append(result_row)
            
            # Создаем DataFrame
            result_df = pd.DataFrame(result_data)
            
            # Рассчитываем ранги ОД
            self.logger.log_debug("Рассчитываем ранги ОД...")
            
            # РАНГ ОД ДЛЯ УРОВНЯ BANK (по всем данным)
            result_df['РАНГ ОД ДЛЯ УРОВНЯ BANK'] = result_df['ОД ТЕКУЩИЙ'].rank(method='min', ascending=False)
            
            # РАНГ ОД ДЛЯ УРОВНЯ TB (по каждому ТБ отдельно)
            result_df['РАНГ ОД ДЛЯ УРОВНЯ TB'] = result_df.groupby('ТБ')['ОД ТЕКУЩИЙ'].rank(method='min', ascending=False)
            
            # Конвертируем ранги в проценты
            total_count = len(result_df)
            result_df['РАНГ ОД ДЛЯ УРОВНЯ BANK'] = (result_df['РАНГ ОД ДЛЯ УРОВНЯ BANK'] / total_count * 100).round(2)
            
            # Конвертируем ранги ТБ в проценты для каждого ТБ
            for tb in result_df['ТБ'].unique():
                tb_mask = result_df['ТБ'] == tb
                tb_count = tb_mask.sum()
                result_df.loc[tb_mask, 'РАНГ ОД ДЛЯ УРОВНЯ TB'] = (
                    result_df.loc[tb_mask, 'РАНГ ОД ДЛЯ УРОВНЯ TB'] / tb_count * 100
                ).round(2)
            
            self.logger.log_debug(f"Обработано данных: {len(result_df)} строк, {len(result_df.columns)} колонок")
            self.logger.log_info(LOG_MESSAGES["processing_end"])
            
            return result_df
            
        except Exception as e:
            error_msg = LOG_MESSAGES["processing_error"].format(str(e))
            self.logger.log_error(error_msg)
            self.logger.log_debug(LOG_MESSAGES["details_error"].format(traceback.format_exc()))
            self.errors_count += 1
            return pd.DataFrame()
    
    def save_outputs(self, processed_data):
        """
        Сохранение обработанных данных в выходные файлы
        
        Args:
            processed_data (pd.DataFrame): Обработанные данные
        """
        if processed_data.empty:
            self.logger.log_error(LOG_MESSAGES["no_data_to_save"])
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
                self.logger.log_debug(LOG_MESSAGES["file_saved_debug_old"].format(file_path))
                self.outputs_created += 1
                
            except Exception as e:
                error_msg = LOG_MESSAGES["save_error"].format(f"{output_config['name']}: {str(e)}")
                self.logger.log_error(error_msg)
                self.logger.log_debug(LOG_MESSAGES["details_error"].format(traceback.format_exc()))
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
            error_msg = LOG_MESSAGES["processing_error"].format(str(e))
            self.logger.log_error(error_msg)
            self.logger.log_debug(LOG_MESSAGES["details_error"].format(traceback.format_exc()))
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
            logger.log_info(LOG_MESSAGES["mode_create_test"])
            generator = TestDataGenerator(WORK_DIR, logger)
            generator.create_sample_data()
            print(LOG_MESSAGES["test_data_success"])
            
        else:
            # Режим обработки данных (по умолчанию)
            logger.log_info(LOG_MESSAGES["mode_process"])
            processor = DataProcessor(WORK_DIR, logger)
            processor.run()
            print(LOG_MESSAGES["process_success"])
        
        # Логируем завершение работы
        logger.log_end()
        
    except Exception as e:
        print(LOG_MESSAGES["main_critical_error"].format(str(e)))
        sys.exit(1)

# =============================================================================
# ТОЧКА ВХОДА
# =============================================================================

if __name__ == "__main__":
    main()
