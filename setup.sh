#!/bin/bash

# Скрипт настройки и запуска программы обработки данных
# Автор: AI Assistant
# Версия: 1.0.0

echo "=== Настройка окружения для программы обработки данных ==="

# Проверяем наличие conda
if ! command -v conda &> /dev/null; then
    echo "ОШИБКА: Conda не установлена. Установите Anaconda или Miniconda."
    exit 1
fi

echo "✓ Conda найдена"

# Создаем новое окружение
echo "Создание conda окружения 'data_processor'..."
conda env create -f environment.yml

if [ $? -eq 0 ]; then
    echo "✓ Окружение создано успешно"
else
    echo "ОШИБКА: Не удалось создать окружение"
    exit 1
fi

# Активируем окружение
echo "Активация окружения..."
source $(conda info --base)/etc/profile.d/conda.sh
conda activate data_processor

if [ $? -eq 0 ]; then
    echo "✓ Окружение активировано"
else
    echo "ОШИБКА: Не удалось активировать окружение"
    exit 1
fi

# Проверяем установленные пакеты
echo "Проверка установленных пакетов..."
python -c "import pandas, openpyxl; print('✓ Все необходимые пакеты установлены')"

# Создаем рабочую структуру папок
echo "Создание рабочей структуры папок..."
WORK_DIR="/Users/orionflash/Desktop/MyProject/Reaction_Effectiv_LOAD/WORK"
mkdir -p "$WORK_DIR/INPUT"
mkdir -p "$WORK_DIR/OUTPUT"
mkdir -p "$WORK_DIR/LOGS"

echo "✓ Рабочие папки созданы в $WORK_DIR"

echo ""
echo "=== Настройка завершена успешно! ==="
echo ""
echo "Для запуска программы выполните:"
echo "  conda activate data_processor"
echo "  python main.py"
echo ""
echo "Поместите входные Excel файлы в папку: $WORK_DIR/INPUT"
echo "Результаты будут сохранены в папку: $WORK_DIR/OUTPUT"
echo "Логи будут сохранены в папку: $WORK_DIR/LOGS"
