import csv
from pathlib import Path
import pandas as pd

# Имя дистрибьютора
this_script = Path(__file__).resolve()
distributor_name = this_script.stem

# Пути
base_dir = Path("C:/Users/user/Desktop/Пример скрипта для подготовки данных к загрузке на python")
registry_path = base_dir / "Реестр файлов" / "registry.csv"
output_dir = base_dir / "Итоговые отчеты"
template_path = base_dir / "Шапка готового отчета" / "Шапка для готового отчета дистрибьютора.xlsx"
output_dir.mkdir(parents=True, exist_ok=True)

# Загрузка шаблона
template_df = pd.read_excel(template_path)
template_columns = template_df.columns.tolist()

# Соответствие колонок
column_mapping = {
    "Дистрибьютор_Филиал": "Региональная компания",
    "Месяц": "Дата",
    "Код SKU": "Код",
    "Номенклатура": "Товар",
    "Получатель_Код": "Код адреса доставки",
    "Получатель_ЮЛ": "Клиент",
    "Получатель_ИНН": "ИНН",
    "Получатель_Факт. адрес": "Адрес доставки",
    "Получатель_Регион": "Регион доставки",
    "Получатель_Город": "Город доставки",
    "Рынок сбыта": "Признак тендер",
    "Отгрузки_УПАК.": "Количество",
    "Дистрибьютор_Название": distributor_name,
    "Файл": None  # заполним вручную
}

# Чтение реестра
with open(registry_path, encoding="utf-8-sig", newline="") as f:
    reader = list(csv.DictReader(f, delimiter=";"))
    fieldnames = reader[0].keys() if reader else []

# Обработка
for row in reader:
    if row["Дистрибьютор"] == distributor_name:
        input_path = Path(row["Путь"])
        try:
            # Определяем лист
            if input_path.suffix.lower() in [".xls", ".xlsx"]:
                xl = pd.ExcelFile(input_path)
                sheet_name = xl.sheet_names[0]

                if len(xl.sheet_names) > 1:
                    for name in xl.sheet_names:
                        df_check = xl.parse(name, nrows=10)
                        if any(
                            all(header in row for header in column_mapping.values() if header)
                            for _, row in df_check.iterrows()
                        ):
                            sheet_name = name
                            break

                df = xl.parse(sheet_name)
            else:
                try:
                    df = pd.read_csv(input_path, sep=";", encoding="utf-8")
                except UnicodeDecodeError:
                    df = pd.read_csv(input_path, sep=";", encoding="cp1251")

            # Подготовка финального DataFrame
            result = pd.DataFrame(columns=template_columns)

            for final_col, source_col in column_mapping.items():
                if source_col in df.columns:
                    result[final_col] = df[source_col]
                elif isinstance(source_col, str):
                    result[final_col] = ""  # на случай отсутствия

            # Специальные колонки
            result["Дистрибьютор_Название"] = distributor_name
            result["Файл"] = str(input_path)

            # Преобразование "Признак тендер"
            if "Рынок сбыта" in result.columns:
                result["Рынок сбыта"] = result["Рынок сбыта"].map({
                    "Да": "Тендер",
                    "Нет": "Коммерция"
                }).fillna(result["Рынок сбыта"])

            # Сохранение результата
            output_path = output_dir / input_path.name
            if input_path.suffix.lower() in [".xls", ".xlsx"]:
                result.to_excel(output_path, index=False)
            else:
                result.to_csv(output_path, sep=";", index=False, encoding="utf-8-sig")

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
