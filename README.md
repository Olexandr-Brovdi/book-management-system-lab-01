# Система обліку книг

Лабораторна робота №1 «Організація професійного Python-проєкту та Git».
Автор: **Бровді Олександр**. Варіант **2**. Курс «Професійний Python».

## Можливості

Додавання книги (назва, автор, рік, сторінки), пошук за автором без урахування
регістру, фільтрація за точним роком, найбільша книга та середня кількість
сторінок. Додатково: меню, введення з клавіатури, валідація, сортування,
dataclass, власний виняток і конфігураційні константи.
Рік: 1–2100; сторінки: додатне ціле число. Порожній каталог: середнє 0.0,
найбільша книга None. За однакової кількості сторінок вибирається перша книга.
Дані зберігаються в пам'яті до завершення програми.
Демонстраційні значення сторінок є навчальними даними для умовних видань.

## Встановлення Python 3.11 або новіший

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -e .
```

Linux/macOS: `source .venv/bin/activate` замість команди активації Windows.

## Запуск

```text
python -m book_manager
python -m book_manager.main
python -m book_manager --interactive
book-manager --interactive
```

Перші дві команди запускають демонстрацію без очікування вводу.
Приклад результату: найбільша книга «Кобзар», 240 стор.; середня кількість
сторінок чотирьох книг — 168.00. Меню починає роботу з порожнього каталогу.

## Тести

```text
python -m unittest discover -s tests -v
```

## Структура

```text
src/book_manager/
    __init__.py, __main__.py, config.py, exceptions.py
    main.py, models.py, services.py
tests/
    test_main.py, test_models.py, test_services.py
tools/
    build_report.py, create_report_assets.py
docs/
    Lab_01_Report_Variant_2.docx
    screenshots/
Lab_01/
    pyproject.toml, README.md
    src/book_manager/ (ідентична копія)
    tests/, tools/
    report/report.docx
```

Ієрархія повторює репозиторій-зразок DioDen5/student-grading-system-lab-01.
Основний код знаходиться в корені; Lab_01 є автономною копією для перевірки.
`models.py` — модель Book; `services.py` — операції над каталогом;
`main.py` — меню і демонстрація; `config.py` — межі; `exceptions.py` — помилки.
Звіт містить опис, код, результати перевірки та відповіді на 30 запитань.
