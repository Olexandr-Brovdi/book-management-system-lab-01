"""Build the report for Brovdi Oleksandr, variant 2, from project files."""
from pathlib import Path
import shutil
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

ROOT = Path(__file__).resolve().parents[1]

QA = [
('Що таке Python-проєкт?', 'Сукупність коду, конфігурації, документації та тестів для розв’язання визначеної задачі.'),
('Чим проєкт відрізняється від одного Python script?', 'Проєкт має узгоджену структуру, метадані, залежності й тести; скрипт зазвичай є одним виконуваним файлом.'),
('Що таке Python module?', 'Окремий файл .py з визначеннями функцій, класів і змінних.'),
('Що таке Python package?', 'Пакет об’єднує модулі в простір імен. Звичайний пакет містить __init__.py; namespace package може не містити цього файла.'),
('Для чого потрібний __init__.py?', 'Позначає звичайний пакет і може задавати його публічний інтерфейс або метадані, наприклад __version__.'),
('Що таке src-layout?', 'Організація, за якої пакети розташовані в каталозі src, окремо від тестів і конфігурації.'),
('Які переваги має src-layout?', 'Запобігає випадковому імпорту коду з кореня та допомагає тестувати саме встановлений пакет.'),
('Для чого використовується virtual environment?', 'Ізолює встановлені пакети та їхні версії від інших проєктів і системного Python.'),
('Як створити .venv?', 'Виконати python -m venv .venv у каталозі проєкту.'),
('Як активувати virtual environment?', 'У PowerShell: .venv\\Scripts\\Activate.ps1; у Linux/macOS: source .venv/bin/activate.'),
('Для чого використовується pip?', 'Для встановлення, оновлення, перегляду та видалення Python-пакетів.'),
('Чому рекомендується python -m pip?', 'Так pip запускається саме тим інтерпретатором Python, який обрано для проєкту.'),
('Для чого потрібний pyproject.toml?', 'Описує систему збирання, метадані проєкту та конфігурацію інструментів.'),
('Що містить секція [project]?', 'Назву, версію, опис, підтримувану версію Python, авторів та інші метадані; підсекції можуть містити команди запуску.'),
('Що таке editable installation?', 'Встановлення через python -m pip install -e .: зміни у вихідному коді доступні без повторного встановлення пакета.'),
('Для чого використовується if __name__ == "__main__"?', 'Запускає точку входу під час виконання модуля, але не під час його імпорту.'),
('Що таке Git?', 'Розподілена система контролю версій для відстеження історії змін і співпраці.'),
('Що таке repository?', 'Сховище файлів проєкту разом з історією Git та службовими даними.'),
('Що таке commit?', 'Зафіксований стан проіндексованих змін з автором, часом, повідомленням та посиланням на батьківські коміти.'),
('Для чого використовується git status?', 'Показує гілку, стан індексу, змінені та невідстежувані файли.'),
('Для чого git add?', 'Додає поточний вміст вибраних файлів до індексу для наступного коміту.'),
('Для чого git commit?', 'Створює коміт зі змін, підготовлених в індексі.'),
('Що показує git log?', 'Історію комітів; --oneline показує скорочений ідентифікатор і повідомлення кожного коміту.'),
('Для чого потрібен .gitignore?', 'Задає шаблони невідстежуваних файлів, які Git має ігнорувати; вже відстежувані файли автоматично не видаляє.'),
('Чому .venv не потрібно комітити?', 'Містить відтворювані залежності та локальні шляхи, займає місце й залежить від платформи.'),
('Для чого потрібний README?', 'Пояснює призначення проєкту, встановлення, запуск, перевірку, структуру та авторство.'),
('Що таке type hints?', 'Анотації очікуваних типів параметрів, результатів і змінних; самі по собі не виконують перевірку під час запуску.'),
('Для чого програму розділяють на modules?', 'Для відокремлення відповідальностей, повторного використання, зрозумілості та незалежного тестування.'),
('Що таке business logic?', 'Правила обробки предметних даних: тут пошук книг, фільтрація, визначення найбільшої книги й середньої кількості сторінок.'),
('Чому потрібно відокремлювати application entry point від business logic?', 'Щоб обчислення не залежали від input/print та могли використовуватися в тестах або іншому інтерфейсі.'),
]


def build() -> None:
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    section.top_margin = section.bottom_margin = Cm(2)
    section.left_margin, section.right_margin = Cm(2.5), Cm(1.5)
    for name in ('Normal', 'Title', 'Heading 1', 'Heading 2'):
        style = doc.styles[name]
        style.font.name = 'Times New Roman'
        style.font.size = Pt(12 if name == 'Normal' else 14)
        style.font.color.rgb = RGBColor(0, 0, 0)
    for border in doc.styles.element.xpath('.//w:pBdr'):
        border.getparent().remove(border)
    doc.styles['Normal'].paragraph_format.space_after = Pt(6)
    doc.styles['Normal'].paragraph_format.line_spacing = 1.12

    def p(text, style=None):
        return doc.add_paragraph(text, style)

    def heading(text):
        doc.add_heading(text, level=1)

    def page(text):
        doc.add_page_break()
        heading(text)

    def code(text, size=9):
        para = p('')
        para.paragraph_format.line_spacing = 1
        run = para.add_run(text.rstrip())
        run.font.name, run.font.size = 'Consolas', Pt(size)

    def file(name, size=9):
        code((ROOT / name).read_text(encoding='utf-8'), size)

    def picture(name, caption):
        doc.add_picture(str(ROOT / 'docs/screenshots' / (name + '.png')), width=Cm(16.5))
        p(caption)

    cover = [
        ('ЛЬВІВСЬКИЙ НАЦІОНАЛЬНИЙ УНІВЕРСИТЕТ\nІМЕНІ ІВАНА ФРАНКА', 14, 65),
        ('ЗВІТ ДО ЛАБОРАТОРНОЇ РОБОТИ №1', 16, 20),
        ('«Організація професійного Python-проєкту та Git»', 14, 18),
        ('з курсу «Професійний Python»\nВаріант №2 — «Система обліку книг»', 14, 45),
        ('Виконав: Бровді Олександр\nГрупа: ФеП-32\nПеревірив(ла): Свелеба Сергій Андрійович\nДата: 07.11.2026', 12, 70),
        ('Львів — 2026', 14, 0),
    ]
    for index, (text, size, after) in enumerate(cover):
        para = p(text, 'Title' if index == 1 else None)
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.space_after = Pt(after)
        for run in para.runs:
            run.font.size = Pt(size)

    page('Мета та постановка задачі')
    p('Мета роботи — навчитися організовувати Python-проєкт за принципом src-layout, працювати з пакетами й модулями, ізольованим середовищем, pyproject.toml, Git і документацією.')
    p('Індивідуальний варіант 2 — система обліку книг. Для кожної книги зберігаються назва, автор, рік видання та кількість сторінок. Реалізовано додавання, пошук за автором, фільтрацію за роком, визначення найбільшої книги й обчислення середньої кількості сторінок.')
    heading('Структура проєкту')
    code('''book-management-system-lab-01/
├── .gitignore, pyproject.toml, README.md
├── src/book_manager/
│   ├── __init__.py, __main__.py
│   ├── config.py, exceptions.py
│   └── models.py, services.py, main.py
├── tests/test_models.py, test_services.py, test_main.py
├── tools/build_report.py, create_report_assets.py
├── docs/Lab_01_Report_Variant_2.docx
│   └── screenshots/
└── Lab_01/
    ├── .gitignore, pyproject.toml, README.md
    ├── src/book_manager/, tests/, tools/
    └── report/report.docx''')
    p('Структуру звірено з DioDen5/student-grading-system-lab-01. Коренева та вкладена копії коду ідентичні. Назву пакета й назву звіту змінено відповідно до варіанта 2.')
    heading('Призначення модулів')
    p('models.py описує dataclass Book і валідацію; services.py містить бізнес-логіку; main.py — консольний інтерфейс; config.py — допустимі межі; exceptions.py — BookValidationError; __init__.py — версію пакета; __main__.py — запуск пакета як модуля.')

    page('Модель даних та перевірка введення')
    file('src/book_manager/models.py', 9)
    p('Порожні назва й автор, недопустимий рік та недодатна кількість сторінок відхиляються. Boolean та дробові числа не приймаються замість цілих. Frozen dataclass не дозволяє змінювати перевірені поля після створення.')
    file('src/book_manager/config.py')
    file('src/book_manager/exceptions.py')

    page('Основні операції над книгами')
    file('src/book_manager/services.py', 9)
    p('Автор порівнюється точно після видалення крайніх пробілів і casefold. Фільтр року використовує точну рівність. Для порожнього списку найбільша книга відсутня, а середнє дорівнює 0.0. Сортування не змінює початковий каталог.')

    page('Точка входу та інтерактивний режим')
    main = (ROOT / 'src/book_manager/main.py').read_text(encoding='utf-8')
    code(main[main.index('def main()'):], 10)
    p('Демонстраційний режим показує всі п’ять операцій без очікування вводу. Параметр --interactive запускає меню з порожнім каталогом. Дані зберігаються в оперативній пам’яті протягом сеансу.')
    code(main[main.index('def run_menu()'):main.index('            elif command == "3":')], 9)
    p('Команди 3–7 викликають пошук за автором, фільтр року, найбільшу книгу, середнє та сортування. Помилки введення обробляються без аварійного завершення; EOF та Ctrl+C завершують сеанс.')

    page('Налаштування та документація')
    file('pyproject.toml', 9)
    p('[build-system] визначає setuptools; [project] — метадані й Python 3.11+; [project.scripts] додає команду book-manager; пошук пакетів налаштовано на src. Секція ruff задає довжину рядка й цільову версію Python.')
    heading('Вміст gitignore')
    file('.gitignore', 9)
    heading('Фрагмент README та встановлення')
    code('python -m venv .venv\n.venv\\Scripts\\Activate.ps1\npython -m pip install -e .\npython -m book_manager\npython -m book_manager --interactive', 9)

    page('Результати запуску застосунку')
    picture('application_output_1', 'Рисунок 1. Зафіксований вивід: додавання, каталог, пошук та рік.')
    picture('application_output_2', 'Рисунок 2. Продовження: найбільша книга, середнє та сортування.')
    p('Демонстраційні обсяги умовних видань: 240, 192, 160 і 80 сторінок. Середнє: (240 + 192 + 160 + 80) / 4 = 168.00. Найбільша книга — «Кобзар», 240 сторінок. За автором Іван Франко знайдено дві книги, за 1840 роком — одну.')

    page('Перевірка та історія Git')
    picture('test_results', 'Рисунок 3. Результат запуску 12 автоматизованих тестів.')
    picture('git_status', 'Рисунок 4. Стан Git під час підготовки звіту. Нові docs/tools ще не додані.')
    picture('git_log', 'Рисунок 5. Історія етапів розробки до додавання звіту.')
    p('Зображення містять реальні результати команд, оформлені для читабельності. Звіт та матеріали перевірки додаються окремим наступним комітом. Тести перевіряють п’ять операцій, порожній каталог, межі даних, сортування, меню та обробку помилок.')

    for start in (0, 10, 20):
        page(f'Відповіді на контрольні запитання {start + 1}–{start + 10}')
        for index, (question, answer) in enumerate(QA[start:start + 10], start + 1):
            para = p('')
            para.paragraph_format.keep_with_next = True
            para.add_run(f'{index}. {question}').bold = True
            p(answer)

    page('Висновок')
    p('Створено консольну систему обліку книг відповідно до варіанта 2. Застосунок має src-layout, окрему модель даних, бізнес-логіку й точку входу, а також type hints, конфігурацію та власний виняток. Перевірено встановлення у virtual environment, запуск пакета й усі основні операції. Додано README, .gitignore та pyproject.toml. Розробку розділено на логічні коміти Git.')
    p('Додатково реалізовано введення з клавіатури, меню, валідацію, сортування та автоматизовані тести. Окремо опрацьовано порожній каталог і неправильне введення. Проєкт придатний для подальшого розширення збереженням даних і новими інтерфейсами.')
    heading('Використані матеріали')
    p('Методичні вказівки Lab_01.docx, розділи 1–29; індивідуальний варіант 2; вимоги до звіту з розділу 27.')
    p('Зразок структури: https://github.com/DioDen5/student-grading-system-lab-01')
    doc.core_properties.author = 'Бровді Олександр'
    doc.core_properties.title = 'Лабораторна робота 1 Система обліку книг Варіант 2'
    output = ROOT / 'docs/Lab_01_Report_Variant_2.docx'
    output.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output)
    target = ROOT / 'Lab_01/report/report.docx'
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(output, target)
    print(output)


if __name__ == '__main__':
    build()
