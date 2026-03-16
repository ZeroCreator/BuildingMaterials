"""
Конфигурация проекта "Поставщики стройматериалов Хабаровск"

Содержит:
- Список материалов для поиска
- Поля для Excel
- Настройки вывода
"""

from pathlib import Path


class Config:
    """Конфигурация проекта"""

    # === Регион ===
    REGION = "Хабаровский край"
    CITY = "Хабаровск"

    # === Список стройматериалов ===
    MATERIALS = [
        {
            "name": "Сваи железобетонные C150.35-13.1",
            "unit": "шт",
            "note": "Бетон B35 W8 F150",
        },
        {
            "name": "Сваи железобетонные С100.35-13.1",
            "unit": "шт",
            "note": "Бетон B35 W8 F150",
        },
        {
            "name": "Бой кирпича",
            "unit": "м³",
            "note": "",
        },
        {
            "name": "Вторичный щебень",
            "unit": "м³",
            "note": "",
        },
        {
            "name": "Геотекстиль 200г/м²",
            "unit": "м²",
            "note": "плотность 200г/м²",
        },
        {
            "name": "Песок строительный",
            "unit": "м³",
            "note": "Мк 2,0-2,5",
        },
    ]

    # === Поля Excel ===
    EXCEL_FIELDS = [
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

    # === Пути ===
    @classmethod
    def get_project_root(cls) -> Path:
        """Корневая директория проекта"""
        return Path(__file__).parent.parent

    @classmethod
    def get_results_dir(cls) -> Path:
        """Директория для результатов"""
        return cls.get_project_root() / "results"

    @classmethod
    def get_materials_list(cls) -> list:
        """Список названий материалов"""
        return [m["name"] for m in cls.MATERIALS]

    @classmethod
    def get_materials_dict(cls) -> dict:
        """Словарь материалов с единицами измерения"""
        return {m["name"]: m["unit"] for m in cls.MATERIALS}


# === Вспомогательные функции ===


def get_latest_file(pattern: str) -> Path | None:
    """Поиск последнего файла по паттерну"""
    import glob
    import os

    results_dir = Config.get_results_dir()
    files = glob.glob(str(results_dir / pattern))

    if not files:
        return None

    files.sort(key=os.path.getmtime, reverse=True)
    return Path(files[0])


def get_previous_version() -> Path | None:
    """Поиск предыдущей версии файла"""
    return get_latest_file("suppliers_*.xlsx")
