from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(r"C:\Users\Нурали\Современные веб технологии")
TEMPLATE = Path(r"C:\Users\Нурали\OneDrive\Desktop\титул.docx")

REPORTS = [
    {
        "file": "Отчет_ЛР1.docx", "number": 1,
        "topic": "Семантический HTML5 и доступная форма",
        "objective": "Создать семантически корректную страницу планировщика задач с формой ввода и средствами доступности.",
        "tools": "HTML5, встроенная проверка полей браузером и семантические элементы страницы.",
        "work": [
            "В файле lab1.html структура разделена на header, nav, main, section и footer. Ссылка перехода к основному содержимому помогает пользователю клавиатуры быстрее попасть к форме.",
            "Форма содержит поля названия, срока, приоритета и описания. Каждое поле связано с подписью через for и id; для обязательных данных заданы required и ограничения длины. Для даты применяется тип date, а подсказка к описанию связана через aria-describedby.",
            "Список-пример размечен элементами ul, li и article, а дата — элементом time с машиночитаемым значением. Атрибуты lang и viewport помогают браузеру корректно определить язык и масштаб страницы."
        ],
        "result": "Подготовлена самостоятельная HTML-страница планировщика с формой, встроенной валидацией и семантической навигацией. Ручная проверка в W3C Validator и оценка Lighthouse Accessibility в рамках этой работы не выполнялись.",
        "conclusion": "В ходе работы применены семантические элементы HTML5, подписи и ограничения полей формы, а также базовые приёмы доступной навигации."
    },
    {
        "file": "Отчет_ЛР2.docx", "number": 2,
        "topic": "Адаптивный интерфейс на CSS3",
        "objective": "Развить страницу планировщика, создав адаптивную раскладку по принципу mobile-first.",
        "tools": "CSS3, CSS Grid, Flexbox, медиазапросы и пользовательские CSS-свойства.",
        "work": [
            "В файле lab2.html сохранена семантическая структура предыдущей работы и добавлено оформление через CSS. Пользовательские свойства в :root задают семейство шрифтов, цвета и фоновые поверхности, а box-sizing: border-box упрощает расчёт размеров.",
            "Основная область использует CSS Grid с двумя колонками: формой и перечнем задач. Flexbox применяется для навигации и панели заголовка, позволяя элементам переноситься при нехватке ширины.",
            "Базовые стили рассчитаны на узкий экран. Медиазапрос при ширине до 700 пикселей перестраивает карточки в одну колонку; для широкого экрана задаются дополнительные отступы. Состояния фокуса выделены контрастной рамкой."
        ],
        "result": "Создана адаптивная версия планировщика с карточками формы и списка задач. Вёрстка предусматривает мобильную, промежуточную и широкую ширину экрана.",
        "conclusion": "Grid, Flexbox, медиазапросы и CSS-переменные позволили адаптировать интерфейс, сохранив читаемость и доступность клавиатурного фокуса."
    },
    {
        "file": "Отчет_ЛР3.docx", "number": 3,
        "topic": "Интерактивность на JavaScript ES6+",
        "objective": "Добавить создание, изменение состояния и удаление задач с обновлением DOM и сохранением данных.",
        "tools": "JavaScript ES6, ES-модуль, DOM API и localStorage браузера.",
        "work": [
            "В lab3.html состояние представлено массивом объектов задач. Для идентификаторов используется crypto.randomUUID; при запуске данные читаются из localStorage с обработкой некорректного JSON и резервным начальным списком.",
            "Обработчик submit получает поля через FormData и создаёт новую задачу. Функция render перестраивает список DOM-элементами, а текст пользовательского ввода помещается через textContent. Обработчики checkbox и кнопки удаления обновляют массив с помощью map и filter.",
            "После изменения список сохраняется в localStorage. Количество активных задач пересчитывается при отображении, а сообщение о добавлении или удалении выводится в live-области для вспомогательных технологий."
        ],
        "result": "Страница позволяет добавлять задачи, отмечать выполненные и удалять записи. Данные остаются в локальном хранилище браузера после перезагрузки страницы.",
        "conclusion": "В работе применены современный синтаксис JavaScript, обработчики событий, безопасное обновление DOM и хранение состояния между посещениями страницы."
    },
    {
        "file": "Отчет_ЛР4.docx", "number": 4,
        "topic": "Асинхронная загрузка данных через REST API",
        "objective": "Получить список задач с удалённого REST API и отобразить состояния загрузки, ошибки и пустой результат.",
        "tools": "Fetch API, Promise, async/await и DummyJSON Todos API.",
        "work": [
            "В lab4.html функция loadTodos выполняет запрос fetch к DummyJSON. До запроса кнопка блокируется, список очищается, а live-область показывает состояние загрузки.",
            "После получения ответа проверяется response.ok, затем разбирается JSON и проверяется, что поле todos является массивом. Для непустого списка создаются элементы с задачей и признаком выполнения; пустой массив отображается отдельным сообщением.",
            "Блок try/catch показывает понятное сообщение при сетевой ошибке, неуспешном HTTP-ответе или неожиданной структуре данных. Блок finally возвращает кнопку в доступное состояние независимо от исхода запроса."
        ],
        "result": "Реализована загрузка демонстрационных задач с API с отдельной индикацией загрузки, результата, пустого списка и ошибки. Работа API зависит от интернет-соединения.",
        "conclusion": "Асинхронный запрос разделён на проверяемые этапы, а все основные состояния запроса представлены пользователю без зависания интерфейса."
    },
    {
        "file": "Отчет_ЛР5.docx", "number": 5,
        "topic": "Планировщик задач на React",
        "objective": "Перенести планировщик на React и организовать интерфейс из повторно используемых компонентов.",
        "tools": "React, компоненты, props, useState, useEffect, useMemo, useContext и React DOM.",
        "work": [
            "В lab5.html приложение разбито на TaskForm, TaskList и TaskItem. Задача передаётся дочернему компоненту через props, а key задаётся по устойчивому идентификатору элемента списка.",
            "TaskProvider хранит массив задач и сообщения интерфейса в состоянии React. Контекст TaskContext предоставляет компонентам операции добавления, переключения выполнения и удаления без передачи обработчиков через каждый промежуточный уровень.",
            "useEffect синхронизирует состояние с localStorage, а useMemo сохраняет объект действий контекста между рендерами. TaskForm управляет значениями полей и отправляет форму без перезагрузки страницы."
        ],
        "result": "Создано React-приложение с компонентным списком задач, общей областью состояния и сохранением данных в браузере. React и React DOM подключаются с CDN, поэтому для запуска требуется интернет и локальный веб-сервер.",
        "conclusion": "Компонентная декомпозиция отделила форму от списка, а Context упростил совместное использование состояния и действий приложения."
    },
    {
        "file": "Отчет_ЛР6.docx", "number": 6,
        "topic": "Маршрутизация и валидация формы в React",
        "objective": "Завершить приложение маршрутами, проверкой пользовательского ввода и понятной обратной связью.",
        "tools": "React, React Router, контролируемые поля формы и браузерное локальное хранилище.",
        "work": [
            "В lab6.html компонент BrowserRouter задаёт клиентскую навигацию. Routes отображает список задач по адресу /, форму по адресу /new и страницу NotFound для неизвестного пути. Ссылки Link и NavLink обеспечивают переходы между разделами.",
            "Форма управляет значениями через useState и проверяет название, наличие срока и то, что срок не находится в прошлом. Ошибки показываются рядом с полями, поля получают aria-invalid и связь с поясняющим текстом через aria-describedby.",
            "Контекст хранит задачи и состояние обратной связи. Успешное сохранение переводит пользователя к списку, где появляется сообщение об успехе; ошибка проверки выводится в alert-области. Данные сохраняются в localStorage."
        ],
        "result": "Получено приложение с отдельными маршрутами списка и создания задачи, проверкой полей, состояниями ошибок и успеха, а также страницей для неизвестного адреса.",
        "conclusion": "Клиентская маршрутизация и доступные сообщения об ошибках завершили последовательное развитие планировщика от HTML-формы до React-приложения."
    },
]

def set_run_font(run, size=14, bold=None):
    run.font.name = "Times New Roman"
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "Times New Roman")
    run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold

def replace_para(para, value):
    if para.runs:
        para.runs[0].text = value
        for run in para.runs[1:]:
            run.text = ""
    else:
        para.add_run(value)

def add_body_para(doc, text, *, heading=False):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT if heading else WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Cm(0) if heading else Cm(1.25)
    p.paragraph_format.space_before = Pt(6 if heading else 0)
    p.paragraph_format.space_after = Pt(0 if not heading else 3)
    p.paragraph_format.line_spacing = 1.5
    if heading:
        p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    set_run_font(run, 14, True if heading else False)
    return p

for item in REPORTS:
    doc = Document(TEMPLATE)
    replace_para(doc.paragraphs[10], f"по лабораторной работе №{item['number']}")
    replace_para(doc.paragraphs[11], f"Тема: «{item['topic']}»")
    # Keep the original title page untouched apart from its two title slots.
    doc.add_page_break()
    add_body_para(doc, "ЦЕЛЬ РАБОТЫ", heading=True)
    add_body_para(doc, item["objective"])
    add_body_para(doc, "ИНСТРУМЕНТЫ И ТЕХНОЛОГИИ", heading=True)
    add_body_para(doc, item["tools"])
    add_body_para(doc, "ХОД РАБОТЫ", heading=True)
    for paragraph in item["work"]:
        add_body_para(doc, paragraph)
    add_body_para(doc, "РЕЗУЛЬТАТ РАБОТЫ", heading=True)
    add_body_para(doc, item["result"])
    add_body_para(doc, "ЗАКЛЮЧЕНИЕ", heading=True)
    add_body_para(doc, item["conclusion"])
    # Set the requested font consistently in newly generated and retained paragraphs.
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if run.text.strip():
                set_run_font(run, 14, run.bold)
    out = ROOT / item["file"]
    doc.save(out)
    print(out)
