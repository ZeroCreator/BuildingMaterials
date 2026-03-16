"""
Оформление данных в Excel
Создаёт XLSX файл из JSON/CSV данных с 3 вкладками:
- Поставщики
- Различия (сравнение с предыдущей версией)
- Статистика
"""

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd


class ExcelFormatter:
    """Оформление данных в Excel с сравнением версий"""

    def __init__(self, results_dir: Path = None):
        if results_dir:
            self.results_dir = Path(results_dir)
        else:
            self.results_dir = Path(__file__).parent.parent / "results"
        self.results_dir.mkdir(parents=True, exist_ok=True)

    def load_data(self, filepath: str) -> list[dict]:
        """Загрузка данных из JSON или CSV"""
        filepath = Path(filepath)

        if not filepath.exists():
            print(f"❌ Файл не найден: {filepath}")
            return []

        if filepath.suffix == ".json":
            with open(filepath, encoding="utf-8") as f:
                return json.load(f)
        elif filepath.suffix == ".csv":
            df = pd.read_csv(filepath, encoding="utf-8")
            return df.to_dict("records")

        print(f"❌ Неподдерживаемый формат: {filepath.suffix}")
        return []

    def find_previous_version(self, current_date: str) -> tuple[Path, list] | None:
        """Поиск предыдущей версии файла"""
        import glob
        import os

        pattern = str(self.results_dir / "suppliers_*.xlsx")
        files = glob.glob(pattern)

        if not files:
            return None

        current_date_obj = datetime.strptime(current_date, "%Y%m%d")

        valid_files = []
        for f in files:
            fname = Path(f).stem
            if "_" in fname:
                try:
                    file_date = fname.split("_")[1]
                    file_date_obj = datetime.strptime(file_date, "%Y%m%d")
                    if file_date_obj < current_date_obj:
                        valid_files.append((f, file_date_obj))
                except ValueError:
                    continue

        if not valid_files:
            return None

        valid_files.sort(key=lambda x: x[1], reverse=True)
        prev_file = valid_files[0][0]

        try:
            df = pd.read_excel(prev_file, sheet_name="Поставщики")
            return Path(prev_file), df.to_dict("records")
        except Exception as e:
            print(f"⚠️ Не удалось загрузить предыдущую версию: {e}")
            return None

    def compare_versions(self, new_data: list, prev_data: list) -> list[dict]:
        """Сравнение с предыдущей версией"""
        if not prev_data:
            return []

        changes = []
        prev_keys = {
            (r.get("Наименование поставщика", ""), r.get("Материал", "")): r
            for r in prev_data
        }
        new_keys = {
            (r.get("Наименование поставщика", ""), r.get("Материал", "")): r
            for r in new_data
        }

        prev_set = set(prev_keys.keys())
        new_set = set(new_keys.keys())

        for key in new_set - prev_set:
            changes.append(
                {
                    "Тип изменения": "➕ Добавлено",
                    "Организация": key[0],
                    "Материал": key[1],
                    "Что изменилось": "Появился в базе",
                    "Дата": datetime.now().strftime("%d.%m.%Y"),
                }
            )

        for key in prev_set - new_set:
            changes.append(
                {
                    "Тип изменения": "➖ Удалено",
                    "Организация": key[0],
                    "Материал": key[1],
                    "Что изменилось": "Нет в новой версии",
                    "Дата": datetime.now().strftime("%d.%m.%Y"),
                }
            )

        for key in prev_set & new_set:
            prev_rec = prev_keys[key]
            new_rec = new_keys[key]

            diffs = []
            for field in ["Телефон", "Email", "Цена", "Адрес"]:
                prev_val = str(prev_rec.get(field, "")).strip()
                new_val = str(new_rec.get(field, "")).strip()
                if prev_val != new_val and (prev_val or new_val):
                    diffs.append(f"{field}: {prev_val} → {new_val}")

            if diffs:
                changes.append(
                    {
                        "Тип изменения": "✏️ Изменено",
                        "Организация": key[0],
                        "Материал": key[1],
                        "Что изменилось": "; ".join(diffs),
                        "Дата": datetime.now().strftime("%d.%m.%Y"),
                    }
                )

        return changes

    def create_stats(self, data: list, prev_data: list = None) -> list[dict]:
        """Создание статистики"""
        stats = []

        stats.append({"Метрика": "ОБЩИЕ ДАННЫЕ", "Значение": ""})
        stats.append({"Метрика": "Всего поставщиков", "Значение": len(data)})

        materials_count = {}
        for r in data:
            mat = r.get("Материал", "Неизвестно")
            materials_count[mat] = materials_count.get(mat, 0) + 1

        stats.append({"Метрика": "", "Значение": ""})
        stats.append({"Метрика": "ПО МАТЕРИАЛАМ", "Значение": ""})
        for mat, count in sorted(
            materials_count.items(), key=lambda x: x[1], reverse=True
        ):
            stats.append({"Метрика": mat, "Значение": count})

        stats.append({"Метрика": "", "Знаzenie": ""})
        stats.append({"Метрика": "ЗАПОЛНЕННОСТЬ ПОЛЕЙ", "Значение": ""})

        for field in ["Телефон", "Email", "ИНН", "Цена", "Адрес"]:
            filled = sum(1 for r in data if r.get(field))
            pct = (filled / len(data) * 100) if data else 0
            stats.append({"Метрика": field, "Значение": f"{filled} ({pct:.1f}%)"})

        if prev_data:
            stats.append({"Метрика": "", "Значение": ""})
            stats.append({"Метрика": "СРАВНЕНИЕ С ПРЕДЫДУЩЕЙ ВЕРСИЕЙ", "Значение": ""})
            stats.append(
                {
                    "Метрика": "Предыдущая версия",
                    "Значение": f"{len(prev_data)} записей",
                }
            )
            stats.append(
                {"Метрика": "Новая версия", "Значение": f"{len(data)} записей"}
            )

        return stats

    def format_excel(self, data: list, date: str = None) -> Path:
        """Основная функция форматирования"""
        if not data:
            print("❌ Нет данных для оформления")
            return None

        if not date:
            date = datetime.now().strftime("%Y%m%d")

        output_file = self.results_dir / f"suppliers_{date}.xlsx"

        prev = self.find_previous_version(date)
        prev_data = prev[1] if prev else None
        prev_filename = prev[0].name if prev else None

        if prev_data:
            changes = self.compare_versions(data, prev_data)
        else:
            changes = []

        stats = self.create_stats(data, prev_data)

        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            df_main = pd.DataFrame(data)
            cols = [
                c
                for c in [
                    "Город",
                    "Наименование поставщика",
                    "ИНН",
                    "Контактное лицо",
                    "Телефон",
                    "Email",
                    "Материал",
                    "Единица",
                    "Цена",
                    "Примечание",
                    "Источник",
                    "Ссылка",
                    "Дата",
                    "Адрес",
                ]
                if c in df_main.columns
            ]
            df_main = df_main[cols]
            df_main.to_excel(writer, sheet_name="Поставщики", index=False)

            if changes:
                df_changes = pd.DataFrame(changes)
                df_changes.to_excel(writer, sheet_name="Различия", index=False)

            df_stats = pd.DataFrame(stats)
            df_stats.to_excel(writer, sheet_name="Статистика", index=False)

        print(f"✅ Создан файл: {output_file}")
        if prev_filename:
            print(f"   📊 Сравнение с: {prev_filename}")
        if changes:
            print(f"   📈 Изменений: {len(changes)}")
        print(f"   📦 Записей: {len(data)}")

        return output_file


def main():
    parser = argparse.ArgumentParser(description="Оформление Excel из JSON/CSV")
    parser.add_argument("input", help="Путь к JSON или CSV файлу")
    parser.add_argument(
        "--date", help="Дата в формате YYYYMMDD (по умолчанию - сегодня)"
    )
    parser.add_argument("--output-dir", help="Директория для результатов")

    args = parser.parse_args()

    formatter = ExcelFormatter(args.output_dir)

    data = formatter.load_data(args.input)
    if not data:
        sys.exit(1)

    result = formatter.format_excel(data, args.date)
    if result:
        print(f"\n🎉 Готово!")
    else:
        sys.exit(1)


if __name__ == "__main__":
    main()
