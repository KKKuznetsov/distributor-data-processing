import csv
from pathlib import Path
import pandas as pd

# Указать вручную имя дистрибьютора
distributor_name = "Катрен"

# Явно задаём базовую директорию
base_dir = Path("C:/Users/user/Desktop/Пример скрипта для подготовки данных к загрузке на python")
registry_path = base_dir / "Реестр файлов" / "registry.csv"
template_path = base_dir / "Шапка готового отчета" / "Шапка для готового отчета дистрибьютора.xlsx"
output_dir = base_dir / "Итоговые отчеты"
output_dir.mkdir(parents=True, exist_ok=True)

# Загрузка шаблона
template_df = pd.read_excel(template_path)
template_columns = template_df.columns.tolist()

# Загрузка реестра
with open(registry_path, encoding="utf-8-sig", newline="") as f:
    reader = list(csv.DictReader(f, delimiter=";"))
    fieldnames = reader[0].keys() if reader else []

# Словарь соответствий
column_map = {
    "Дистрибьютор_Филиал": "Филиал",
    "Получатель_ЮЛ": "Клиент",
    "Получатель_Регион": "Регион",
    "Получатель_Город": "Город",
    "Получатель_Факт. адрес": "Улица",
    "Номенклатура": "Товар",
    "Получатель_ИНН": "ИНН клиента",
    "Код SKU": "UID товара",
    "Месяц": "День",
    "Рынок сбыта": "Аптека.РУ",
    "Отгрузки_УПАК.": "Продажи, шт.",
    "Дистрибьютор_Название": None,
    "Файл": None
}

# Обработка
for row in reader:
    if row["Дистрибьютор"] == distributor_name:
        input_path = Path(row["Путь"])
        try:
            # Чтение исходного файла
            if input_path.suffix.lower() in [".xls", ".xlsx"]:
                df = pd.read_excel(input_path)
            elif input_path.suffix.lower() == ".csv":
                try:
                    df = pd.read_csv(input_path, sep=";", encoding="utf-8")
                except UnicodeDecodeError:
                    df = pd.read_csv(input_path, sep=";", encoding="cp1251")

            # Подготовка финальной таблицы
            output_df = pd.DataFrame(columns=template_columns)

            for target_col, source_col in column_map.items():
                if source_col is None:
                    if target_col == "Дистрибьютор_Название":
                        output_df[target_col] = distributor_name
                    elif target_col == "Файл":
                        output_df[target_col] = str(input_path)
                elif source_col in df.columns:
                    if target_col == "Рынок сбыта":
                        output_df[target_col] = df[source_col].apply(
                            lambda x: "Аптека.ру" if str(x).strip().lower() == "да" else "Коммерция"
                        )
                    else:
                        output_df[target_col] = df[source_col]
                else:
                    output_df[target_col] = None

            # Сохраняем файл
            output_path = output_dir / input_path.name
            if input_path.suffix.lower() in [".xls", ".xlsx"]:
                output_df.to_excel(output_path, index=False)
            else:
                output_df.to_csv(output_path, index=False, sep=";", encoding="utf-8-sig")

            row["Статус"] = "Обработан"
            print(f"[OK] Обработан: {input_path.name}")

        except Exception as e:
            row["Статус"] = "Не обработан"
            print(f"[Ошибка] {input_path.name}: {e}")

# Обновление реестра
with open(registry_path, mode="w", encoding="utf-8-sig", newline="") as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=";")
    writer.writeheader()
    writer.writerows(reader)

print(f"[INFO] Статусы обновлены в {registry_path}")
