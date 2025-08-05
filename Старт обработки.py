import os
import csv
import subprocess
from pathlib import Path

# Базовая директория
base_dir = Path("C:/Users/user/Desktop/Пример скрипта для подготовки данных к загрузке на python")

# Пути
distributors_dir = base_dir / "Дистрибьюторы"
scripts_dir = base_dir / "Скрипты обработки для конкретных дистрибьюторов"
output_folder = base_dir / "Реестр файлов"
output_file = output_folder / "registry.csv"

# Создание выходной папки
output_folder.mkdir(parents=True, exist_ok=True)

rows = []

# === ШАГ 1: Сканирование папки и формирование списка записей ===
for path in distributors_dir.rglob("*.*"):
    if path.suffix.lower() in [".xlsx", ".xls", ".csv"]:
        try:
            distributor = path.parts[-4]
            period = path.parts[-3]
            purchase_type = path.parts[-2]
        except IndexError:
            distributor = period = purchase_type = "Не определено"

        template_script = scripts_dir / f"{distributor}.py"
        status = ""  # Пока не заполняем

        rows.append([
            str(path),
            distributor,
            period,
            purchase_type,
            str(template_script) if template_script.exists() else "",
            status
        ])

# === ШАГ 2: Сохраняем registry.csv ===
with open(output_file, mode="w", newline="", encoding="utf-8-sig") as f:
    writer = csv.writer(f, delimiter=";", quoting=csv.QUOTE_MINIMAL)
    writer.writerow(["Путь", "Дистрибьютор", "Период", "Тип закупки", "Шаблон", "Статус"])
    writer.writerows(rows)

print(f"✅ Реестр успешно сохранён: {output_file}")

# === ШАГ 3: Запуск всех уникальных скриптов обработки ===
launched = set()
for row in rows:
    script_path = row[4]
    if script_path and script_path not in launched:
        if Path(script_path).exists():
            try:
                result = subprocess.run(
                    ["python", script_path],
                    check=True,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    timeout=120
                )
                print(f"🚀 Скрипт выполнен: {Path(script_path).name}")
                print(result.stdout.decode("utf-8", errors="ignore"))
            except subprocess.CalledProcessError as e:
                print(f"❌ Ошибка в скрипте {Path(script_path).name}:")
                print(e.stderr.decode("utf-8", errors="ignore"))
            launched.add(script_path)
