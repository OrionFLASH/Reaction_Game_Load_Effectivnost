#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Временный скрипт для проверки формата ТН 10 в созданных файлах
"""

import pandas as pd
import os

# Путь к файлам
input_dir = "/Users/orionflash/Desktop/MyProject/Reaction_Effectiv_LOAD/WORK/INPUT"
file1 = "data1_20250822_163010.xlsx"
file2 = "data2_20250822_163010.xlsx"

print("=== ПРОВЕРКА ФОРМАТА ТН 10 ===")

# Проверяем файл 1
if os.path.exists(os.path.join(input_dir, file1)):
    print(f"\nФайл 1: {file1}")
    df1 = pd.read_excel(os.path.join(input_dir, file1))
    
    print(f"Количество строк: {len(df1)}")
    print(f"Колонки: {list(df1.columns)}")
    
    # Проверяем первые 10 ТН
    print("\nПервые 10 ТН 10:")
    for i, tn in enumerate(df1['ТН 10'].head(10)):
        print(f"  {i+1}: '{tn}' (тип: {type(tn)}, длина: {len(str(tn))})")
    
    # Проверяем, есть ли ТН с лидирующими нулями
    tn_with_zeros = [tn for tn in df1['ТН 10'] if str(tn).replace('TN_', '').startswith('0')]
    print(f"\nТН с лидирующими нулями: {len(tn_with_zeros)}")
    if tn_with_zeros:
        print("Примеры:")
        for tn in tn_with_zeros[:5]:
            print(f"  '{tn}'")
    
    # Проверяем минимальную и максимальную длину ТН
    tn_lengths = [len(str(tn)) for tn in df1['ТН 10']]
    print(f"Длина ТН: мин={min(tn_lengths)}, макс={max(tn_lengths)}")
    
    # Проверяем количество лидирующих нулей
    tn_leading_zeros = []
    for tn in df1['ТН 10']:
        tn_str = str(tn).replace('TN_', '')
        leading_zeros = len(tn_str) - len(tn_str.lstrip('0'))
        tn_leading_zeros.append(leading_zeros)
    
    print(f"Лидирующие нули: мин={min(tn_leading_zeros)}, макс={max(tn_leading_zeros)}")
    print(f"ТН с 1+ лидирующими нулями: {sum(1 for x in tn_leading_zeros if x > 0)}")
    print(f"ТН с 2+ лидирующими нулями: {sum(1 for x in tn_leading_zeros if x >= 2)}")
    print(f"ТН с 3+ лидирующими нулями: {sum(1 for x in tn_leading_zeros if x >= 3)}")
    
    # Примеры ТН с разным количеством лидирующих нулей
    print("\nПримеры ТН с лидирующими нулями:")
    for i, tn in enumerate(df1['ТН 10'].head(20)):
        tn_str = str(tn).replace('TN_', '')
        leading_zeros = len(tn_str) - len(tn_str.lstrip('0'))
        if leading_zeros > 0:
            print(f"  '{tn}' -> {leading_zeros} лидирующих нулей")
    
else:
    print(f"Файл {file1} не найден")

# Проверяем файл 2
if os.path.exists(os.path.join(input_dir, file2)):
    print(f"\nФайл 2: {file2}")
    df2 = pd.read_excel(os.path.join(input_dir, file2))
    
    print(f"Количество строк: {len(df2)}")
    
    # Проверяем первые 10 ТН
    print("\nПервые 10 ТН 10:")
    for i, tn in enumerate(df2['ТН 10'].head(10)):
        print(f"  {i+1}: '{tn}' (тип: {type(tn)}, длина: {len(str(tn))})")
    
    # Проверяем, есть ли ТН с лидирующими нулями
    tn_with_zeros = [tn for tn in df2['ТН 10'] if str(tn).replace('TN_', '').startswith('0')]
    print(f"\nТН с лидирующими нулями: {len(tn_with_zeros)}")
    if tn_with_zeros:
        print("Примеры:")
        for tn in tn_with_zeros[:5]:
            print(f"  '{tn}'")
    
    # Проверяем минимальную и максимальную длину ТН
    tn_lengths = [len(str(tn)) for tn in df2['ТН 10']]
    print(f"Длина ТН: мин={min(tn_lengths)}, макс={max(tn_lengths)}")
    
    # Проверяем количество лидирующих нулей
    tn_leading_zeros = []
    for tn in df2['ТН 10']:
        tn_str = str(tn).replace('TN_', '')
        leading_zeros = len(tn_str) - len(tn_str.lstrip('0'))
        tn_leading_zeros.append(leading_zeros)
    
    print(f"Лидирующие нули: мин={min(tn_leading_zeros)}, макс={max(tn_leading_zeros)}")
    print(f"ТН с 1+ лидирующими нулями: {sum(1 for x in tn_leading_zeros if x > 0)}")
    print(f"ТН с 2+ лидирующими нулями: {sum(1 for x in tn_leading_zeros if x >= 2)}")
    print(f"ТН с 3+ лидирующими нулями: {sum(1 for x in tn_leading_zeros if x >= 3)}")
    
    # Примеры ТН с разным количеством лидирующих нулей
    print("\nПримеры ТН с лидирующими нулями:")
    for i, tn in enumerate(df2['ТН 10'].head(20)):
        tn_str = str(tn).replace('TN_', '')
        leading_zeros = len(tn_str) - len(tn_str.lstrip('0'))
        if leading_zeros > 0:
            print(f"  '{tn}' -> {leading_zeros} лидирующих нулей")
    
else:
    print(f"Файл {file2} не найден")

print("\n=== КОНЕЦ ПРОВЕРКИ ===")
