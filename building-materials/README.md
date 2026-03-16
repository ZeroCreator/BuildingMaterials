# Поставщики стройматериалов (Хабаровск)

Сбор и оформление базы поставщиков стройматериалов в г. Хабаровск.

## Workflow

1. **AI собирает данные** → `results/suppliers_YYYYMMDD.json`
2. **Оформление в Excel** → `results/suppliers_YYYYMMDD.xlsx`

## Материалы

- Сваи железобетонные C150.35-13.1
- Сваи железобетонные С100.35-13.1
- Бой кирпича
- Вторичный щебень
- Геотекстиль 200г/м²
- Песок строительный

## Использование

### 1. Запуск AI

Используй промт из `prompts/khabarovsk_suppliers_search.md`

### 2. Оформление Excel

```bash
cd building-materials
pip install -r requirements.txt

python -m src.formatter results/suppliers_20260318.json
```

### 3. С указанием даты

```bash
python -m src.formatter results/suppliers_20260318.json --date 20260318
```

## Результат

Файл `results/suppliers_YYYYMMDD.xlsx` с 3 вкладками:

- **Поставщики** — основные данные
- **Различия** — сравнение с предыдущей версией
- **Статистика** — метрики по материалам и полям

## Структура

```
building-materials/
├── prompts/          # Промты для AI
├── src/              # Python код
│   ├── config.py
│   └── formatter.py
├── results/          # Входные/выходные файлы
└── requirements.txt
```
