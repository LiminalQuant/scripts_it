# OS ⇄ IT Merge Tool

Streamlit-приложение для объединения ведомости ОС с набором IT-файлов
по инвентарному номеру.

## Что делает

1. Сканирует папку с IT-файлами (.xlsx)
2. Показывает реальные уникальные названия колонок
3. Позволяет выбрать:
   - ключ в целевой ведомости
   - ключ в IT-файлах
   - любые колонки для добавления (1+)
4. Делает объединение
5. Формирует Excel с листами:
   - MATCHED — итоговая таблица
   - UNMATCHED — строки из IT-файлов без совпадений

## Запуск локально

```bash
git clone https://github.com/USERNAME/os-it-merge.git
cd os-it-merge
pip install -r requirements.txt
streamlit run app.py
# scripts_it
