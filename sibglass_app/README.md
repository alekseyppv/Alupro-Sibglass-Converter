# Alupro → SibGlass Converter (Desktop)

Production-ready desktop приложение на **Python 3.11+** для автоматизации создания заявки СибГласс на основе файла AluPro.

## 1. Описание проекта

Приложение помогает инженеру/менеджеру:

- выбрать Excel-файл AluPro;
- выбрать шаблон заявки СибГласс;
- автоматически извлечь формулы заполнений;
- сопоставить их с итоговой формулой стеклопакета;
- заполнить заявку с расчетом площадей;
- сохранить результат без ручного копирования данных между файлами.

Поддерживается архитектура **MVC + Services + Repositories** с разделением ответственности и удобством расширения.

## 2. Скриншоты

> Ниже размещены заглушки для GitHub. Замените на реальные изображения после первого запуска.

Скриншоты намеренно не включены в репозиторий в рамках требований к PR (без бинарных `*.png`).

## 3. Архитектура проекта (MVC + Services + Repositories)

### MVC

- **Models** — типизированные доменные сущности (`FormulaItem`, `OrderItem`, `GlassCatalog`).
- **Views** — UI на PySide6 (`MainWindow`, таблица формул, диалоги ручного ввода).
- **Controllers** — orchestration UI-событий и прикладного workflow (`MainController`).

### Services

Бизнес-логика вынесена в сервисы:

- разбор AluPro,
- построение формул,
- валидация,
- запись заявки,
- автосохранение,
- управление справочником стекол.

### Repositories

Изолируют доступ к внешним источникам:

- Excel (pandas/openpyxl),
- `glass.txt`.

## 4. Технологический стек

| Категория | Технологии |
|---|---|
| Язык | Python 3.11+ |
| GUI | PySide6 |
| Excel | pandas, openpyxl |
| Логирование | logging (`errors.log`) |
| Конфиг | JSON (`settings.json`) |
| Сборка | PyInstaller |

## 5. Инструкция по запуску

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install pyside6 pandas openpyxl xlrd
python -m sibglass_app.main
```
> Для чтения старых `.xls` файлов требуется `xlrd>=2.0.1`.

## 6. Сборка в .exe (PyInstaller)

```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed --name SibglassConverter sibglass_app/main.py
```

Результат: `dist/SibglassConverter.exe`.

> При необходимости добавьте `--add-data "sibglass_app/data;sibglass_app/data"` для упаковки начальных данных.

## 7. Структура проекта

```text
sibglass_app/
│
├─ main.py
├─ app.py
│
├─ config/
│   ├─ settings.py
│   └─ paths.py
│
├─ models/
│   ├─ formula_item.py
│   ├─ order_item.py
│   └─ glass_catalog.py
│
├─ views/
│   ├─ main_window.py
│   ├─ dialogs.py
│   └─ formula_table.py
│
├─ controllers/
│   └─ main_controller.py
│
├─ services/
│   ├─ alupro_parser.py
│   ├─ sibglass_writer.py
│   ├─ formula_builder.py
│   ├─ glass_catalog_service.py
│   ├─ autosave_service.py
│   └─ validation_service.py
│
├─ repositories/
│   ├─ excel_repository.py
│   └─ glass_file_repository.py
│
├─ utils/
│   ├─ logger.py
│   ├─ excel_utils.py
│   └─ text_utils.py
│
├─ data/
│   └─ (runtime files: glass.txt, autosave.tmp)
│
└─ README.md
```

## 8. Пример `glass.txt`

Файл `glass.txt` создается при первом сохранении/ручном добавлении значений. Пример структуры:

```text
Стекло наружное:
...

Стекло среднее:
...

Стекло внутреннее:
...

Рамки:
...
```

## 9. Пример workflow

1. Выбрать файл AluPro (`.xlsx/.xls`) — валидация по маркеру `Заполнения`.
2. Выбрать файл заявки СибГласс — валидация по маркеру `ЗАЯВКА НА РАСЧЕТ СТЕКЛОПАКЕТОВ`.
3. Заполнить поля "Заказчик" и "Адрес".
4. Настроить стекла/рамки и флаги `Зак`/`Арг`.
5. Проверить таблицу найденных формул, при необходимости вручную поправить итоговую формулу.
6. Нажать **Сохранить заявку**.
7. Приложение заполнит таблицу заявки, рассчитает площади и сохранит Excel-файл.

## 10. Лицензия

Проект распространяется по лицензии **MIT**.

См. файл [LICENSE](LICENSE).
