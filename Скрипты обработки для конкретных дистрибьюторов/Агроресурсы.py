import csv
from pathlib import Path
import pandas as pd
from dateutil import parser

# Название дистрибьютора по имени файла скрипта
distributor_name = Path(__file__).stem

# Пути
base_dir = Path("C:/Users/user/Desktop/Пример скрипта для подготовки данных к загрузке на python")
registry_path = base_dir / "Реестр файлов" / "registry.csv"
output_dir = base_dir / "Итоговые отчеты"
template_path = base_dir / "Шапка готового отчета" / "Шапка для готового отчета дистрибьютора.xlsx"

output_dir.mkdir(parents=True, exist_ok=True)

# Загрузка шаблона (порядок столбцов)
template_df = pd.read_excel(template_path)
template_columns = template_df.columns.tolist()

# Маппинг русских месяцев
month_map = {
    'янв': 'Jan', 'фев': 'Feb', 'мар': 'Mar', 'апр': 'Apr',
    'май': 'May', 'июн': 'Jun', 'июл': 'Jul', 'авг': 'Aug',
    'сен': 'Sep', 'окт': 'Oct', 'ноя': 'Nov', 'дек': 'Dec'
}

def convert_date_russian(date_str):
    try:
        if pd.isna(date_str):
            return pd.NaT
        for ru, en in month_map.items():
            if ru in str(date_str).lower():
                date_str = str(date_str).lower().replace(ru, en)
                break
        return parser.parse(date_str)
    except Exception:
        return pd.NaT

# Загрузка реестра
with open(registry_path, encoding="utf-8-sig", newline="") as f:
    reader = list(csv.DictReader(f, delimiter=";"))
    fieldnames = reader[0].keys() if reader else []

# Обработка
for row in reader:
    if row["Дистрибьютор"] == distributor_name:
        input_path = Path(row["Путь"])
        try:
            # Поиск шапки по первым 10 строкам
            for header_row in range(10):
                df_try = pd.read_excel(input_path, header=header_row)
                normalized_cols = [str(col).strip().lower() for col in df_try.columns]
                if "количество" in normalized_cols:
                    df = df_try.copy()
                    break
            else:
                raise ValueError("Не найдена строка с колонкой 'количество' в первых 10 строках.")

            # Фильтрация по количеству
            df = df[df["количество"] != 0]

            # Обработка даты
            df["дата_документа"] = df["дата_документа"].apply(convert_date_russian)
            df["Месяц"] = df["дата_документа"].dt.strftime("%d.%m.%Y")
            df["Накладная"] = df["номер_документа"].astype(str) + " от " + df["дата_документа"].dt.strftime("%d.%m.%Y")

            # Словарь соответствий
            column_mapping = {
                "Получатель_ИНН": "инн",
                "Получатель_ЮЛ": "клиент",
                "Получатель_Регион": "область_район",
                "Получатель_Город": "город",
                "Получатель_Факт. адрес": "адрес",
                "Номенклатура": "название",
                "Отгрузки_УПАК.": "количество",
                "Дистрибьютор_Филиал": "филиал",
                "Получатель_Код": "код_клиента",
                "Накладная": "Накладная",
                "Код SKU": "код_товара",
                "Месяц": "Месяц"
            }

            # Сбор итогового датафрейма
            final_df = pd.DataFrame(columns=template_columns)
            for target_col, source_col in column_mapping.items():
                final_df[target_col] = df.get(source_col, "")

            final_df["Дистрибьютор_Название"] = row["Дистрибьютор"]
            final_df["Файл"] = str(input_path)

            # Сохранение
            output_path = output_dir / input_path.name
            if input_path.suffix.lower() in [".xls", ".xlsx"]:
                final_df.to_excel(output_path, index=False)
            else:
                final_df.to_csv(output_path, index=False, sep=";", encoding="utf-8-sig")

            print(f"[OK] Обработан: {input_path.name}")
            row["Статус"] = "Обработан"

        except Exception as e:
            print(f"[ОШИБКА] {input_path.name}: {e}")
            row["Статус"] = "Не обработан"

# Сохраняем обновлённый реестр
with open(registry_path, mode="w", encoding="utf-8-sig", newline="") as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=";")
    writer.writeheader()
    writer.writerows(reader)

print(f"[INFO] Статусы обновлены в {registry_path}")
