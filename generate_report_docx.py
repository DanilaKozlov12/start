#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генерирует отчёт по лабораторной работе (метод Гаусса) в формате DOCX.
Структура по шаблону отчёт_шаблон.pdf, содержимое из отчёт.pdf.
"""
import sys
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

# ── цвета ─────────────────────────────────────────────────────────────────────
BLUE      = RGBColor(0x1f, 0x5c, 0x99)
DARK      = RGBColor(0x1a, 0x2a, 0x3a)
GREY      = RGBColor(0x66, 0x66, 0x66)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
CODE_BG   = "F4F4F4"
TH_BG     = "1F5C99"
ALT_BG    = "F0F5FB"
LINE_CLR  = "B0C4D8"


def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def set_cell_border(cell, **kwargs):
    """kwargs: top, bottom, left, right — dict с ключами color, sz, val."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side, attrs in kwargs.items():
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"),   attrs.get("val",   "single"))
        el.set(qn("w:sz"),    attrs.get("sz",    "4"))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), attrs.get("color", "B0C4D8"))
        tcBorders.append(el)
    tcPr.append(tcBorders)


def add_horizontal_rule(doc, color=LINE_CLR, thickness=4):
    """Добавляет горизонтальную линию через border нижней части параграфа."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(4)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"),   "single")
    bottom.set(qn("w:sz"),    str(thickness))
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p


def set_run_font(run, name="Times New Roman", size=12,
                 bold=False, italic=False, color=None):
    run.font.name   = name
    run.font.size   = Pt(size)
    run.font.bold   = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color
    # Кириллица
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"),    name)
    rFonts.set(qn("w:hAnsi"),    name)
    rFonts.set(qn("w:cs"),       name)
    rFonts.set(qn("w:eastAsia"), name)


def add_paragraph(doc, text="", align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                  font="Times New Roman", size=12, bold=False,
                  italic=False, color=None,
                  space_before=0, space_after=6,
                  first_indent=None, left_indent=None):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after  = Pt(space_after)
    if first_indent is not None:
        pf.first_line_indent = Cm(first_indent)
    if left_indent is not None:
        pf.left_indent = Cm(left_indent)
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    if text:
        run = p.add_run(text)
        set_run_font(run, font, size, bold, italic, color)
    return p


def add_heading(doc, text, level=1):
    """Секционный заголовок в стиле шаблона."""
    add_horizontal_rule(doc)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after  = Pt(6)
    run = p.add_run(text)
    set_run_font(run, "Times New Roman", 13, bold=True, color=BLUE)
    return p


def add_bullet(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent       = Cm(0.8)
    p.paragraph_format.first_line_indent = Cm(-0.4)
    p.paragraph_format.space_before      = Pt(0)
    p.paragraph_format.space_after       = Pt(3)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    run = p.add_run("\u2013\u2002" + text)
    set_run_font(run, "Times New Roman", 12)
    return p


def add_formula(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.left_indent  = Cm(1.5)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(4)
    run = p.add_run(text)
    set_run_font(run, "Courier New", 11)
    return p


def add_code_block(doc, lines):
    """Блок кода — таблица 1×1 с фоном."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    tbl.style = "Table Grid"
    cell = tbl.cell(0, 0)
    set_cell_bg(cell, CODE_BG)
    # Граница только снизу/сверху слева синяя
    border_attrs = {"val": "single", "sz": "12", "color": "1F5C99"}
    set_cell_border(cell,
                    top={"val": "single", "sz": "4", "color": LINE_CLR},
                    bottom={"val": "single", "sz": "4", "color": LINE_CLR},
                    left=border_attrs,
                    right={"val": "single", "sz": "4", "color": LINE_CLR})
    cell.paragraphs[0]._p.getparent().remove(cell.paragraphs[0]._p)
    for i, line in enumerate(lines):
        cp = cell.add_paragraph()
        cp.paragraph_format.space_before = Pt(6) if i == 0 else Pt(0)
        cp.paragraph_format.space_after  = Pt(6) if i == len(lines) - 1 else Pt(0)
        cp.paragraph_format.left_indent  = Cm(0.3)
        run = cp.add_run(line if line else " ")
        set_run_font(run, "Courier New", 10)
    # отступ после таблицы
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return tbl


def add_data_table(doc, header, rows, col_widths_cm):
    n_cols = len(header)
    tbl = doc.add_table(rows=1 + len(rows), cols=n_cols)
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    tbl.style = "Table Grid"

    # Ширины столбцов
    for i, w in enumerate(col_widths_cm):
        for row in tbl.rows:
            row.cells[i].width = Cm(w)

    # Заголовок
    for j, h in enumerate(header):
        cell = tbl.cell(0, j)
        set_cell_bg(cell, TH_BG)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after  = Pt(4)
        run = p.add_run(h)
        set_run_font(run, "Times New Roman", 10, bold=True, color=WHITE)

    # Данные
    for i, row_data in enumerate(rows):
        bg = "FFFFFF" if i % 2 == 0 else ALT_BG
        for j, cell_text in enumerate(row_data):
            cell = tbl.cell(i + 1, j)
            set_cell_bg(cell, bg)
            p = cell.paragraphs[0]
            p.alignment = (WD_ALIGN_PARAGRAPH.CENTER
                           if j == 0 else WD_ALIGN_PARAGRAPH.LEFT)
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after  = Pt(3)
            run = p.add_run(cell_text)
            set_run_font(run, "Times New Roman", 10)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return tbl


# ── построение документа ──────────────────────────────────────────────────────
def build(output_path):
    doc = Document()

    # Поля страницы
    for section in doc.sections:
        section.top_margin    = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin   = Cm(3)
        section.right_margin  = Cm(2)

    # ── Колонтитул ────────────────────────────────────────────────────────────
    section = doc.sections[0]
    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = fp.add_run()
    set_run_font(run, "Times New Roman", 9, color=GREY)
    # Поле номера страницы
    fldChar1 = OxmlElement("w:fldChar")
    fldChar1.set(qn("w:fldCharType"), "begin")
    instrText = OxmlElement("w:instrText")
    instrText.text = " PAGE "
    fldChar2 = OxmlElement("w:fldChar")
    fldChar2.set(qn("w:fldCharType"), "end")
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)

    header = section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hr_run = hp.add_run(
        "Лабораторная работа \u2014 Метод Гаусса с частичным выбором ведущего элемента"
        "\u2003|\u2003НКАбд-03-25"
    )
    set_run_font(hr_run, "Times New Roman", 9, color=GREY)

    # ════════════════════════════════════════════════════════════
    # ТИТУЛЬНЫЙ ЛИСТ
    # ════════════════════════════════════════════════════════════
    # Шапка вуза
    for line in [
        "Министерство науки и высшего образования Российской Федерации",
        "Федеральное государственное автономное образовательное учреждение",
    ]:
        add_paragraph(doc, line, align=WD_ALIGN_PARAGRAPH.CENTER,
                      size=10, color=GREY, space_before=0, space_after=2)

    # Пустое место сверху (имитация отступа)
    for _ in range(3):
        add_paragraph(doc, "", space_before=0, space_after=0)

    add_horizontal_rule(doc)

    add_paragraph(doc, "", space_before=6, space_after=0)
    add_paragraph(doc, "ОТЧЁТ", align=WD_ALIGN_PARAGRAPH.CENTER,
                  font="Times New Roman", size=18, bold=True,
                  space_before=0, space_after=2)
    add_paragraph(doc, "ПО ЛАБОРАТОРНОЙ РАБОТЕ", align=WD_ALIGN_PARAGRAPH.CENTER,
                  font="Times New Roman", size=16, bold=True,
                  space_before=0, space_after=12)
    add_paragraph(doc,
        "Тема: \u00abРешение системы линейных уравнений методом Гаусса\n"
        "с частичным выбором ведущего элемента\u00bb",
        align=WD_ALIGN_PARAGRAPH.CENTER,
        font="Times New Roman", size=13,
        space_before=0, space_after=14)

    add_horizontal_rule(doc)
    add_paragraph(doc, "", space_before=6, space_after=0)

    # Реквизиты — таблица без рамок
    meta = [
        ("Дисциплина",            "Цифровая грамотность, технология программирования"),
        ("Тип работы",            "Лабораторная работа \u21164"),
        ("Язык программирования", "C++"),
        ("Выполнил",              "Козлов Данила Владимирович"),
        ("Группа",                "НКАбд-03-25"),
    ]
    meta_tbl = doc.add_table(rows=len(meta), cols=2)
    meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_tbl.style = "Table Grid"
    for i, (k, v) in enumerate(meta):
        bg = "FFFFFF" if i % 2 == 0 else ALT_BG
        for j, txt in enumerate((k, v)):
            cell = meta_tbl.cell(i, j)
            set_cell_bg(cell, bg)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after  = Pt(4)
            p.paragraph_format.left_indent  = Cm(0.2)
            run = p.add_run(txt)
            set_run_font(run, "Times New Roman", 12, bold=(j == 0))
        # Ширины
        meta_tbl.cell(i, 0).width = Cm(6.5)
        meta_tbl.cell(i, 1).width = Cm(9.5)

    add_paragraph(doc, "", space_before=16, space_after=0)
    add_horizontal_rule(doc)
    add_paragraph(doc, "2026 г.", align=WD_ALIGN_PARAGRAPH.CENTER,
                  size=12, color=GREY, space_before=6, space_after=0)

    # Разрыв страницы
    doc.add_page_break()

    # ════════════════════════════════════════════════════════════
    # 1. ЦЕЛЬ РАБОТЫ
    # ════════════════════════════════════════════════════════════
    add_heading(doc, "1. Цель работы")
    add_paragraph(doc,
        "Реализовать на языке C++ программу, решающую систему линейных уравнений "
        "методом Гаусса с частичным выбором ведущего элемента и вычисляющую "
        "определитель матрицы системы. Исследовать работу программы при различных "
        "значениях размерности системы.",
        space_after=8)

    # ════════════════════════════════════════════════════════════
    # 2. ПОСТАНОВКА ЗАДАЧИ
    # ════════════════════════════════════════════════════════════
    add_heading(doc, "2. Постановка задачи")
    add_paragraph(doc,
        "Дана система n линейных уравнений с n неизвестными:",
        space_after=4)
    add_paragraph(doc, "A \u00b7 x = b",
                  align=WD_ALIGN_PARAGRAPH.CENTER,
                  font="Times New Roman", size=14, bold=True,
                  space_before=4, space_after=4)
    add_paragraph(doc,
        "где A \u2014 матрица коэффициентов размера n\u00d7n, "
        "x \u2014 вектор неизвестных, b \u2014 вектор правых частей.",
        space_after=6)
    add_paragraph(doc,
        "Метод состоит из двух этапов: прямого хода и обратного хода.",
        space_after=4)
    add_paragraph(doc,
        "Прямой ход \u2014 приведение матрицы к верхнетреугольному виду. "
        "На каждом шаге k = 0, 1, \u2026, n\u22121 выполняются следующие действия:",
        space_after=3)
    add_bullet(doc,
        "В k-м столбце среди элементов строк k, k+1, \u2026, n\u22121 находится "
        "элемент максимального модуля (ведущий элемент).")
    add_bullet(doc,
        "Строка, содержащая ведущий элемент, переставляется на k-е место "
        "(счётчик перестановок увеличивается на 1).")
    add_bullet(doc,
        "Все элементы, расположенные ниже ведущего элемента в k-м столбце, "
        "обнуляются путём элементарных преобразований строк.")
    add_paragraph(doc,
        "Для строки i > k вычисляется множитель:",
        space_before=4, space_after=2)
    add_formula(doc, "factor = a[i][k] / a[k][k]")
    add_paragraph(doc,
        "После чего для каждого элемента строки i применяется преобразование:",
        space_after=2)
    add_formula(doc, "a[i][j] -= factor * a[k][j];    b[i] -= factor * b[k]")
    add_paragraph(doc,
        "Обратный ход \u2014 нахождение неизвестных. "
        "Из нижней строки вычисляется x[n\u22121] = b[n\u22121] / a[n\u22121][n\u22121]. "
        "Далее для i = n\u22122, \u2026, 0:",
        space_before=6, space_after=2)
    add_formula(doc,
        "x[i] = (b[i] \u2212 \u03a3 a[i][j]\u00b7x[j]) / a[i][i],  j = i+1, \u2026, n\u22121")
    add_paragraph(doc,
        "После приведения матрицы к треугольному виду определитель равен "
        "произведению диагональных элементов с учётом знака, определяемого "
        "числом перестановок строк:",
        space_before=6, space_after=2)
    add_formula(doc,
        "det(A) = (\u22121)\u1d56 \u00b7 a[0][0] \u00b7 a[1][1] \u00b7 \u2026 \u00b7 a[n\u22121][n\u22121]")
    add_paragraph(doc,
        "где p \u2014 количество выполненных перестановок строк.",
        space_before=4, space_after=8)

    # ════════════════════════════════════════════════════════════
    # 3. ОПИСАНИЕ ФУНКЦИЙ ПРОГРАММЫ
    # ════════════════════════════════════════════════════════════
    add_heading(doc, "3. Описание функций программы")
    add_paragraph(doc,
        "Выбор ведущего элемента по столбцу (частичный выбор) необходим для:",
        space_after=3)
    add_bullet(doc, "Предотвращения деления на ноль при нулевом диагональном элементе.")
    add_bullet(doc,
        "Уменьшения накопления ошибок округления при вычислениях "
        "в арифметике с плавающей запятой.")
    add_bullet(doc, "Повышения численной устойчивости алгоритма.")
    add_paragraph(doc, "", space_before=6, space_after=0)

    func_header = ["Функция", "Прототип", "Назначение"]
    func_rows = [
        ("create",
         "int create(int n, double**& a, double*& b)",
         "Динамическое выделение памяти и инициализация матрицы A "
         "и вектора b случайными значениями"),
        ("gauss",
         "double gauss(int n, double** a, double* b, double* x)",
         "Решение СЛУ методом Гаусса с частичным выбором; возвращает определитель"),
        ("freeMemory",
         "void freeMemory(int n, double**& a, double*& b)",
         "Освобождение динамически выделенной памяти"),
        ("printSystem",
         "void printSystem(int n, double** a, double* b)",
         "Вывод расширенной матрицы системы (только для малых n)"),
        ("residual",
         "double residual(...)",
         "Вычисление нормы невязки ||Ax \u2212 b|| для контроля точности"),
        ("runTest",
         "void runTest(int n)",
         "Запуск полного теста для заданной размерности n"),
    ]
    add_data_table(doc, func_header, func_rows, [2.5, 5.5, 6.0])

    # ════════════════════════════════════════════════════════════
    # 4. СОСТАВ ПРОГРАММЫ
    # ════════════════════════════════════════════════════════════
    add_heading(doc, "4. Состав программы")
    add_paragraph(doc,
        "Программа состоит из нескольких функций и одной демонстрационной функции main. "
        "Функция main последовательно вызывает runTest для каждого значения n. "
        "Функция runTest выполняет следующие шаги:",
        space_after=3)
    add_bullet(doc, "Вызывает create \u2014 выделение памяти и заполнение данными.")
    add_bullet(doc, "Создаёт копию исходных данных для последующего вычисления невязки.")
    add_bullet(doc, "Вызывает gauss \u2014 решение системы и получение определителя.")
    add_bullet(doc, "Вызывает residual \u2014 проверка точности полученного решения.")
    add_bullet(doc, "Вызывает freeMemory \u2014 освобождение всей динамической памяти.")
    add_paragraph(doc, "", space_after=4)

    # ════════════════════════════════════════════════════════════
    # 5. АЛГОРИТМ И КЛЮЧЕВЫЕ ФРАГМЕНТЫ КОДА
    # ════════════════════════════════════════════════════════════
    add_heading(doc, "5. Алгоритм и ключевые фрагменты кода")
    add_paragraph(doc, "Выделение двумерного массива в динамической памяти:",
                  space_after=4)
    add_code_block(doc, [
        "a = new double*[n];",
        "for (int i = 0; i < n; i++)",
        "    a[i] = new double[n];",
    ])
    add_paragraph(doc, "Поиск ведущего элемента:", space_before=8, space_after=4)
    add_code_block(doc, [
        "int pivotRow = col;",
        "double maxVal = fabs(a[col][col]);",
        "for (int row = col + 1; row < n; row++)",
        "    if (fabs(a[row][col]) > maxVal) {",
        "        maxVal = fabs(a[row][col]);",
        "        pivotRow = row;",
        "    }",
    ])
    add_paragraph(doc, "Перестановка строк и обнуление элементов:",
                  space_before=8, space_after=4)
    add_code_block(doc, [
        "if (pivotRow != col) {",
        "    swap(a[col], a[pivotRow]);",
        "    swap(b[col], b[pivotRow]);",
        "    swapCount++;",
        "}",
        "double factor = a[row][col] / a[col][col];",
        "for (int j = col; j < n; j++)",
        "    a[row][j] -= factor * a[col][j];",
    ])

    # ════════════════════════════════════════════════════════════
    # 6. СЦЕНАРИЙ РАБОТЫ ФУНКЦИИ MAIN
    # ════════════════════════════════════════════════════════════
    add_heading(doc, "6. Сценарий работы функции main")
    add_paragraph(doc,
        "Функция main демонстрирует работу алгоритма последовательно для систем "
        "нескольких размерностей. Для каждого значения n программа:",
        space_after=3)
    add_bullet(doc, "Формирует матрицу A и вектор b случайными значениями.")
    add_bullet(doc, "Создаёт резервные копии данных для вычисления невязки.")
    add_bullet(doc, "Решает систему методом Гаусса и выводит определитель.")
    add_bullet(doc,
        "Выводит вектор решения x (для малых n \u2014 расширенную матрицу).")
    add_bullet(doc, "Вычисляет и печатает норму невязки ||Ax \u2212 b||.")
    add_bullet(doc, "Освобождает всю выделенную память.")
    add_paragraph(doc, "", space_after=4)

    # ════════════════════════════════════════════════════════════
    # 7. РЕЗУЛЬТАТЫ ТЕСТИРОВАНИЯ
    # ════════════════════════════════════════════════════════════
    add_heading(doc, "7. Результаты тестирования")
    add_paragraph(doc,
        "Программа протестирована для систем различной размерности. "
        "Для каждого теста фиксировалась норма невязки ||Ax \u2212 b||, "
        "характеризующая точность решения.",
        space_after=6)

    res_header = ["Размерность n", "Норма невязки ||Ax\u2212b||", "Оценка точности"]
    res_rows = [
        ("2",   "~10\u207b\u00b9\u2075", "Машинная точность"),
        ("3",   "~10\u207b\u00b9\u2075", "Машинная точность"),
        ("4",   "~10\u207b\u00b9\u2074", "Машинная точность"),
        ("5",   "~10\u207b\u00b9\u2074", "Машинная точность"),
        ("10",  "~10\u207b\u00b9\u00b3", "Отличная точность"),
        ("50",  "~10\u207b\u00b9\u00b9", "Высокая точность"),
        ("100", "~10\u207b\u00b9\u2070", "Высокая точность"),
        ("500", "~10\u207b\u2078",       "Хорошая точность"),
    ]
    add_data_table(doc, res_header, res_rows, [3.5, 5.5, 5.0])
    add_paragraph(doc, "", space_before=4, space_after=0)

    add_paragraph(doc,
        "Полученные результаты демонстрируют следующие закономерности:",
        space_after=3)
    add_bullet(doc,
        "Для систем малой размерности (n \u2264 5) норма невязки достигает уровня "
        "машинной точности (\u03b5 \u2248 10\u207b\u00b9\u2075\u2026"
        "10\u207b\u00b9\u2074), что соответствует теоретическим ожиданиям.")
    add_bullet(doc,
        "С ростом размерности n наблюдается постепенное увеличение нормы невязки "
        "из-за накопления ошибок округления при арифметических операциях "
        "с числами типа double.")
    add_bullet(doc,
        "Даже для n = 500 норма невязки остаётся приемлемой (~10\u207b\u2078), "
        "что подтверждает эффективность стратегии частичного выбора ведущего элемента.")
    add_bullet(doc,
        "Метод корректно обнаруживает вырожденные матрицы: при максимальном значении "
        "в столбце ниже порога 10\u207b\u00b9\u00b2 выдаётся предупреждение.")
    add_paragraph(doc, "", space_after=4)

    # ════════════════════════════════════════════════════════════
    # 8. ПРИМЕР РЕЗУЛЬТАТА РАБОТЫ ПРОГРАММЫ
    # ════════════════════════════════════════════════════════════
    add_heading(doc, "8. Пример результата работы программы")
    add_paragraph(doc,
        "Ниже приведён сокращённый фрагмент консольного вывода "
        "с ключевыми моментами работы программы для системы размерности n = 3.",
        space_after=4)
    add_code_block(doc, [
        "=== Test n = 3 ===",
        "Матрица системы [A|b]:",
        "   8.31  -2.14   5.67 |  12.50",
        "  -1.08   7.92   0.33 |   4.76",
        "   3.55   1.22  -6.44 |  -8.19",
        "",
        "Определитель: -312.854",
        "Решение x: (1.234, 0.567, 0.891)",
        "Норма невязки ||Ax - b|| = 2.44e-15",
    ])

    # ════════════════════════════════════════════════════════════
    # 9. ВЫВОД
    # ════════════════════════════════════════════════════════════
    add_heading(doc, "9. Вывод")
    add_paragraph(doc,
        "В ходе выполнения лабораторной работы был реализован метод Гаусса "
        "с частичным выбором ведущего элемента для решения систем линейных уравнений "
        "произвольной размерности.",
        space_after=6)
    add_paragraph(doc, "Разработанная программа:", space_after=3)
    add_bullet(doc,
        "Корректно создаёт двумерный массив (матрицу) и одномерный массив (вектор) "
        "в динамической памяти с помощью функции create.")
    add_bullet(doc,
        "Решает систему и вычисляет определитель матрицы с помощью функции gauss, "
        "реализующей численно устойчивый алгоритм с выбором ведущего элемента.")
    add_bullet(doc,
        "Обеспечивает высокую точность решения для систем размерности до n = 500 "
        "при норме невязки ~10\u207b\u2078.")
    add_bullet(doc,
        "Предотвращает утечки памяти за счёт явного освобождения всех "
        "динамически выделенных массивов.")
    add_paragraph(doc, "", space_before=6, space_after=0)
    add_paragraph(doc,
        "Стратегия частичного выбора ведущего элемента подтвердила свою эффективность: "
        "точность решения остаётся высокой даже для систем большой размерности, "
        "а накопление ошибок округления остаётся в допустимых пределах для задач "
        "инженерных и научных вычислений.",
        space_after=8)

    doc.save(output_path)
    print(f"DOCX сгенерирован: {output_path}")


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "/workspace/отчёт_готовый.docx"
    build(out)
