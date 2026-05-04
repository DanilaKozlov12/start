#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор отчёта по лабораторной работе — метод Гаусса.
Структура по шаблону отчёт_шаблон.pdf, содержимое из отчёт.pdf.
"""
import sys
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageTemplate,
    Paragraph, Spacer, Table, TableStyle,
    PageBreak, KeepTogether,
)
from reportlab.platypus.flowables import HRFlowable
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as pdfcanvas

# ── шрифты ────────────────────────────────────────────────────────────────────
BASE = "/usr/share/fonts/truetype/dejavu/"
pdfmetrics.registerFont(TTFont("R",  BASE + "DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("B",  BASE + "DejaVuSans-Bold.ttf"))
pdfmetrics.registerFont(TTFont("SR", BASE + "DejaVuSerif.ttf"))
pdfmetrics.registerFont(TTFont("SB", BASE + "DejaVuSerif-Bold.ttf"))
pdfmetrics.registerFont(TTFont("M",  BASE + "DejaVuSansMono.ttf"))
pdfmetrics.registerFont(TTFont("MB", BASE + "DejaVuSansMono-Bold.ttf"))

# ── размеры страницы ──────────────────────────────────────────────────────────
PW, PH = A4
ML, MR, MT, MB = 3*cm, 2*cm, 2*cm, 2*cm
TW = PW - ML - MR   # ширина текстового блока

# ── цвета ─────────────────────────────────────────────────────────────────────
C_DARK  = colors.HexColor("#1a2a3a")
C_BLUE  = colors.HexColor("#1f5c99")
C_LBLUE = colors.HexColor("#d6e8f7")
C_ALTBG = colors.HexColor("#f0f5fb")
C_GREY  = colors.HexColor("#666666")
C_LINE  = colors.HexColor("#b0c4d8")
C_WHITE = colors.white
C_CODE  = colors.HexColor("#1e1e1e")
C_CODEBG= colors.HexColor("#f4f4f4")


def S(name, **kw):
    """Создать стиль с базовыми параметрами."""
    defaults = dict(fontName="R", fontSize=11, leading=17,
                    textColor=C_DARK, spaceAfter=0, spaceBefore=0)
    defaults.update(kw)
    return ParagraphStyle(name, **defaults)


# ── стили ─────────────────────────────────────────────────────────────────────
sBody = S("body", alignment=TA_JUSTIFY, spaceAfter=5)
sBullet = S("bullet", leftIndent=16, firstLineIndent=0, spaceAfter=3)
sH1   = S("h1", fontName="SB", fontSize=13, leading=20, textColor=C_BLUE,
           spaceBefore=14, spaceAfter=6)
sMeta_key = S("mk", fontName="B", fontSize=11, leading=15)
sMeta_val = S("mv", fontName="R", fontSize=11, leading=15)
sCode = S("code", fontName="M", fontSize=9, leading=13,
          textColor=C_CODE, leftIndent=0, spaceAfter=0)
sTH   = S("th", fontName="B", fontSize=10, leading=13,
          alignment=TA_CENTER, textColor=C_WHITE)
sTD   = S("td", fontName="R", fontSize=10, leading=14, alignment=TA_LEFT)
sTDc  = S("tdc", fontName="R", fontSize=10, leading=14, alignment=TA_CENTER)
sCoverHead = S("ch", fontName="SB", fontSize=16, leading=24,
               alignment=TA_CENTER, textColor=C_DARK)
sCoverSub  = S("cs", fontName="SR", fontSize=13, leading=20,
               alignment=TA_CENTER, textColor=C_DARK)
sCoverYear = S("cy", fontName="SR", fontSize=12, leading=18,
               alignment=TA_CENTER, textColor=C_GREY)


def sp(h=6):
    return Spacer(1, h)


def hr():
    return HRFlowable(width="100%", thickness=0.6, color=C_LINE,
                      spaceAfter=4, spaceBefore=4)


def heading(text):
    return KeepTogether([hr(), Paragraph(text, sH1)])


def body(text):
    return Paragraph(text, sBody)


def bullet(text):
    return Paragraph("\u2013\u2002" + text, sBullet)


def code_block(lines):
    """lines: list of str — каждая строка кода."""
    rows = [[Paragraph(ln.replace(" ", "\u00a0"), sCode)] for ln in lines]
    t = Table(rows, colWidths=[TW])
    t.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1, -1), C_CODEBG),
        ("LEFTPADDING",  (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING",   (0, 0), (0, 0), 8),
        ("BOTTOMPADDING",(0, -1),(-1, -1), 8),
        ("TOPPADDING",   (0, 1), (-1, -1), 1),
        ("BOTTOMPADDING",(0, 0), (-1, -2), 1),
        ("BOX",          (0, 0), (-1, -1), 0.5, C_LINE),
        ("LINEAFTER",    (0, 0), (0, -1), 3, C_BLUE),
    ]))
    return t


def data_table(header_row, data_rows, col_widths):
    """Таблица данных с заголовком."""
    all_rows = [[Paragraph(h, sTH) for h in header_row]]
    for row in data_rows:
        styles = [sTDc if i == 0 else sTD for i in range(len(row))]
        all_rows.append([Paragraph(cell, styles[i]) for i, cell in enumerate(row)])

    t = Table(all_rows, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0),  C_BLUE),
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [C_WHITE, C_ALTBG]),
        ("GRID",          (0, 0), (-1, -1), 0.4, C_LINE),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 7),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 7),
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    return t


# ── колонтитулы ───────────────────────────────────────────────────────────────
def make_page_templates(doc):
    """Создаёт шаблоны страниц: обложка и основные страницы."""
    cover_frame = Frame(ML, MB, TW, PH - MT - MB, id="cover")
    main_frame  = Frame(ML, MB + 1.2*cm, TW, PH - MT - MB - 1.2*cm, id="main")

    def draw_cover(c, d):
        pass  # без колонтитулов на обложке

    def draw_main(c, d):
        c.saveState()
        # Верхняя линия
        c.setStrokeColor(C_LINE)
        c.setLineWidth(0.5)
        c.line(ML, PH - MT + 4*mm, PW - MR, PH - MT + 4*mm)
        # Колонтитул верх
        c.setFont("R", 8)
        c.setFillColor(C_GREY)
        c.drawString(ML, PH - MT + 6*mm,
                     "Лабораторная работа — Метод Гаусса с частичным выбором ведущего элемента")
        c.drawRightString(PW - MR, PH - MT + 6*mm, "НКАбд-03-25")
        # Нижняя линия + номер страницы
        c.line(ML, MB - 6*mm, PW - MR, MB - 6*mm)
        c.setFont("R", 9)
        c.setFillColor(C_GREY)
        page_str = f"— {d.page - 1} —"
        c.drawCentredString(PW / 2, MB - 11*mm, page_str)
        c.restoreState()

    cover_tpl = PageTemplate(id="Cover", frames=[cover_frame],
                             onPage=draw_cover)
    main_tpl  = PageTemplate(id="Main",  frames=[main_frame],
                             onPage=draw_main)
    doc.addPageTemplates([cover_tpl, main_tpl])


# ── построение документа ──────────────────────────────────────────────────────
def build(output_path):
    doc = BaseDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=ML, rightMargin=MR,
        topMargin=MT,  bottomMargin=MB,
        title="Отчёт по лабораторной работе — Метод Гаусса",
        author="Козлов Данила Владимирович",
    )
    make_page_templates(doc)

    story = []

    # ════════════════════════════════════════════════════════════
    # ТИТУЛЬНЫЙ ЛИСТ
    # ════════════════════════════════════════════════════════════
    story.append(sp(70))

    # Шапка вуза (как в шаблоне)
    story.append(Paragraph(
        "Министерство науки и высшего образования Российской Федерации",
        S("uni", fontName="R", fontSize=10, leading=14, alignment=TA_CENTER, textColor=C_GREY)
    ))
    story.append(sp(2))
    story.append(Paragraph(
        "Федеральное государственное автономное образовательное учреждение",
        S("uni2", fontName="R", fontSize=10, leading=14, alignment=TA_CENTER, textColor=C_GREY)
    ))
    story.append(sp(14))

    story.append(hr())
    story.append(sp(14))

    story.append(Paragraph("ОТЧЁТ", sCoverHead))
    story.append(sp(4))
    story.append(Paragraph("ПО ЛАБОРАТОРНОЙ РАБОТЕ", sCoverHead))
    story.append(sp(18))
    story.append(Paragraph(
        "Тема: «Решение системы линейных уравнений методом Гаусса\n"
        "с частичным выбором ведущего элемента»",
        sCoverSub,
    ))
    story.append(sp(30))
    story.append(hr())
    story.append(sp(14))

    # Реквизиты — таблица без рамок
    meta = [
        ("Дисциплина",            "Цифровая грамотность, технология программирования"),
        ("Тип работы",            "Лабораторная работа \u2116\u00a04"),
        ("Язык программирования", "C++"),
        ("Выполнил",              "Козлов Данила Владимирович"),
        ("Группа",                "НКАбд-03-25"),
    ]
    meta_rows = [[Paragraph(k, sMeta_key), Paragraph(v, sMeta_val)] for k, v in meta]
    meta_tbl = Table(meta_rows, colWidths=[TW * 0.42, TW * 0.58])
    meta_tbl.setStyle(TableStyle([
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LINEBELOW",     (0, 0), (-1, -1), 0.3, C_LINE),
    ]))
    story.append(meta_tbl)
    story.append(sp(40))
    story.append(hr())
    story.append(sp(10))
    story.append(Paragraph("2026 г.", sCoverYear))

    # Переход на основной шаблон
    story.append(PageBreak())
    from reportlab.platypus import NextPageTemplate
    story.insert(story.index(next(f for f in story if isinstance(f, PageBreak))),
                 NextPageTemplate("Main"))

    # ════════════════════════════════════════════════════════════
    # 1. ЦЕЛЬ РАБОТЫ
    # ════════════════════════════════════════════════════════════
    story.append(heading("1. Цель работы"))
    story.append(sp(4))
    story.append(body(
        "Реализовать на языке C++ программу, решающую систему линейных уравнений "
        "методом Гаусса с частичным выбором ведущего элемента и вычисляющую "
        "определитель матрицы системы. Исследовать работу программы при различных "
        "значениях размерности системы."
    ))
    story.append(sp(8))

    # ════════════════════════════════════════════════════════════
    # 2. ПОСТАНОВКА ЗАДАЧИ
    # ════════════════════════════════════════════════════════════
    story.append(heading("2. Постановка задачи"))
    story.append(sp(4))
    story.append(body(
        "Дана система n линейных уравнений с n неизвестными:"
    ))
    story.append(sp(6))
    story.append(Paragraph(
        "A \u00b7 x = b",
        S("formula", fontName="SB", fontSize=13, leading=20, alignment=TA_CENTER)
    ))
    story.append(sp(6))
    story.append(body(
        "где A \u2014 матрица коэффициентов размера n\u00d7n, "
        "x \u2014 вектор неизвестных, "
        "b \u2014 вектор правых частей."
    ))
    story.append(sp(6))
    story.append(body(
        "Метод состоит из двух этапов: прямого хода и обратного хода."
    ))
    story.append(sp(6))
    story.append(body(
        "Прямой ход \u2014 приведение матрицы к верхнетреугольному виду. "
        "На каждом шаге k = 0, 1, \u2026, n\u22121 выполняются следующие действия:"
    ))
    story.append(sp(3))
    story.append(bullet(
        "В k-м столбце среди элементов строк k, k+1, \u2026, n\u22121 находится элемент "
        "максимального модуля (ведущий элемент)."
    ))
    story.append(bullet(
        "Строка, содержащая ведущий элемент, переставляется на k-е место "
        "(счётчик перестановок увеличивается на 1)."
    ))
    story.append(bullet(
        "Все элементы, расположенные ниже ведущего элемента в k-м столбце, "
        "обнуляются путём элементарных преобразований строк."
    ))
    story.append(sp(6))
    story.append(body(
        "Для строки i > k вычисляется множитель:"
    ))
    story.append(sp(4))
    story.append(Paragraph(
        "factor = a[i][k] / a[k][k]",
        S("f2", fontName="M", fontSize=10, leading=14,
          leftIndent=30, alignment=TA_LEFT)
    ))
    story.append(sp(4))
    story.append(body(
        "После чего для каждого элемента строки i применяется преобразование:"
    ))
    story.append(sp(4))
    story.append(Paragraph(
        "a[i][j] -= factor * a[k][j];\u2003\u2003b[i] -= factor * b[k]",
        S("f3", fontName="M", fontSize=10, leading=14,
          leftIndent=30, alignment=TA_LEFT)
    ))
    story.append(sp(8))
    story.append(body(
        "Обратный ход \u2014 нахождение неизвестных. "
        "Из нижней строки вычисляется x[n\u22121] = b[n\u22121] / a[n\u22121][n\u22121]. "
        "Далее для i = n\u22122, \u2026, 0:"
    ))
    story.append(sp(4))
    story.append(Paragraph(
        "x[i] = (b[i] \u2212 \u03a3 a[i][j]\u00b7x[j]) / a[i][i],\u2003j = i+1, \u2026, n\u22121",
        S("f4", fontName="M", fontSize=10, leading=14,
          leftIndent=30, alignment=TA_LEFT)
    ))
    story.append(sp(8))
    story.append(body(
        "После приведения матрицы к треугольному виду определитель равен "
        "произведению диагональных элементов с учётом знака, определяемого "
        "числом перестановок строк:"
    ))
    story.append(sp(4))
    story.append(Paragraph(
        "det(A) = (\u22121)\u1d56 \u00b7 a[0][0] \u00b7 a[1][1] \u00b7 \u2026 \u00b7 a[n\u22121][n\u22121]",
        S("f5", fontName="M", fontSize=10, leading=14,
          leftIndent=30, alignment=TA_LEFT)
    ))
    story.append(sp(4))
    story.append(body("где p \u2014 количество выполненных перестановок строк."))
    story.append(sp(8))

    # ════════════════════════════════════════════════════════════
    # 3. ОПИСАНИЕ ФУНКЦИЙ ПРОГРАММЫ
    # ════════════════════════════════════════════════════════════
    story.append(heading("3. Описание функций программы"))
    story.append(sp(4))
    story.append(body(
        "Выбор ведущего элемента по столбцу (частичный выбор) необходим для:"
    ))
    story.append(sp(3))
    story.append(bullet("Предотвращения деления на ноль при нулевом диагональном элементе."))
    story.append(bullet(
        "Уменьшения накопления ошибок округления при вычислениях "
        "в арифметике с плавающей запятой."
    ))
    story.append(bullet("Повышения численной устойчивости алгоритма."))
    story.append(sp(8))

    func_header = ["Функция", "Прототип", "Назначение"]
    func_rows = [
        ("create",      "int create(int n,\ndouble**& a, double*& b)",
         "Динамическое выделение памяти и инициализация матрицы A "
         "и вектора b случайными значениями"),
        ("gauss",       "double gauss(int n,\ndouble** a, double* b, double* x)",
         "Решение СЛУ методом Гаусса с частичным выбором; возвращает определитель"),
        ("freeMemory",  "void freeMemory(int n,\ndouble**& a, double*& b)",
         "Освобождение динамически выделенной памяти"),
        ("printSystem", "void printSystem(int n,\ndouble** a, double* b)",
         "Вывод расширенной матрицы системы (только для малых n)"),
        ("residual",    "double residual(...)",
         "Вычисление нормы невязки ||Ax \u2212 b|| для контроля точности"),
        ("runTest",     "void runTest(int n)",
         "Запуск полного теста для заданной размерности n"),
    ]
    story.append(data_table(
        func_header,
        [(r[0], r[1], r[2]) for r in func_rows],
        [TW * 0.17, TW * 0.38, TW * 0.45],
    ))
    story.append(sp(8))

    # ════════════════════════════════════════════════════════════
    # 4. СОСТАВ ПРОГРАММЫ
    # ════════════════════════════════════════════════════════════
    story.append(heading("4. Состав программы"))
    story.append(sp(4))
    story.append(body(
        "Программа состоит из нескольких функций и одной демонстрационной функции main. "
        "Функция main последовательно вызывает runTest для каждого значения n. "
        "Функция runTest выполняет следующие шаги:"
    ))
    story.append(sp(3))
    story.append(bullet("Вызывает create \u2014 выделение памяти и заполнение данными."))
    story.append(bullet("Создаёт копию исходных данных для последующего вычисления невязки."))
    story.append(bullet("Вызывает gauss \u2014 решение системы и получение определителя."))
    story.append(bullet("Вызывает residual \u2014 проверка точности полученного решения."))
    story.append(bullet("Вызывает freeMemory \u2014 освобождение всей динамической памяти."))
    story.append(sp(8))

    # ════════════════════════════════════════════════════════════
    # 5. АЛГОРИТМ И КЛЮЧЕВЫЕ ФРАГМЕНТЫ КОДА
    # ════════════════════════════════════════════════════════════
    story.append(heading("5. Алгоритм и ключевые фрагменты кода"))
    story.append(sp(4))
    story.append(body("Выделение двумерного массива в динамической памяти:"))
    story.append(sp(4))
    story.append(code_block([
        "a = new double*[n];",
        "for (int i = 0; i < n; i++)",
        "    a[i] = new double[n];",
    ]))
    story.append(sp(10))
    story.append(body("Поиск ведущего элемента:"))
    story.append(sp(4))
    story.append(code_block([
        "int pivotRow = col;",
        "double maxVal = fabs(a[col][col]);",
        "for (int row = col + 1; row < n; row++)",
        "    if (fabs(a[row][col]) > maxVal) {",
        "        maxVal = fabs(a[row][col]);",
        "        pivotRow = row;",
        "    }",
    ]))
    story.append(sp(10))
    story.append(body("Перестановка строк и обнуление элементов:"))
    story.append(sp(4))
    story.append(code_block([
        "if (pivotRow != col) {",
        "    swap(a[col], a[pivotRow]);",
        "    swap(b[col], b[pivotRow]);",
        "    swapCount++;",
        "}",
        "double factor = a[row][col] / a[col][col];",
        "for (int j = col; j < n; j++)",
        "    a[row][j] -= factor * a[col][j];",
    ]))
    story.append(sp(8))

    # ════════════════════════════════════════════════════════════
    # 6. СЦЕНАРИЙ РАБОТЫ ФУНКЦИИ MAIN
    # ════════════════════════════════════════════════════════════
    story.append(heading("6. Сценарий работы функции main"))
    story.append(sp(4))
    story.append(body(
        "Функция main демонстрирует работу алгоритма последовательно "
        "для систем нескольких размерностей. Для каждого значения n программа:"
    ))
    story.append(sp(3))
    story.append(bullet("Формирует матрицу A и вектор b случайными значениями."))
    story.append(bullet("Создаёт резервные копии данных для вычисления невязки."))
    story.append(bullet("Решает систему методом Гаусса и выводит определитель."))
    story.append(bullet("Выводит вектор решения x (для малых n \u2014 расширенную матрицу)."))
    story.append(bullet("Вычисляет и печатает норму невязки ||Ax \u2212 b||."))
    story.append(bullet("Освобождает всю выделенную память."))
    story.append(sp(8))

    # ════════════════════════════════════════════════════════════
    # 7. РЕЗУЛЬТАТЫ ТЕСТИРОВАНИЯ
    # ════════════════════════════════════════════════════════════
    story.append(heading("7. Результаты тестирования"))
    story.append(sp(4))
    story.append(body(
        "Программа протестирована для систем различной размерности. "
        "Для каждого теста фиксировалась норма невязки ||Ax \u2212 b||, "
        "характеризующая точность решения."
    ))
    story.append(sp(8))

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
    story.append(data_table(
        res_header, res_rows,
        [TW * 0.22, TW * 0.38, TW * 0.40],
    ))
    story.append(sp(10))

    story.append(body("Полученные результаты демонстрируют следующие закономерности:"))
    story.append(sp(3))
    story.append(bullet(
        "Для систем малой размерности (n \u2264 5) норма невязки достигает уровня "
        "машинной точности (\u03b5 \u2248 10\u207b\u00b9\u2075\u2026"
        "10\u207b\u00b9\u2074), что соответствует теоретическим ожиданиям."
    ))
    story.append(bullet(
        "С ростом размерности n наблюдается постепенное увеличение нормы невязки "
        "из-за накопления ошибок округления при арифметических операциях "
        "с числами типа double."
    ))
    story.append(bullet(
        "Даже для n = 500 норма невязки остаётся приемлемой (~10\u207b\u2078), "
        "что подтверждает эффективность стратегии частичного выбора ведущего элемента."
    ))
    story.append(bullet(
        "Метод корректно обнаруживает вырожденные матрицы: при максимальном значении "
        "в столбце ниже порога 10\u207b\u00b9\u00b2 выдаётся предупреждение."
    ))
    story.append(sp(8))

    # ════════════════════════════════════════════════════════════
    # 8. ПРИМЕР РЕЗУЛЬТАТА РАБОТЫ ПРОГРАММЫ
    # ════════════════════════════════════════════════════════════
    story.append(heading("8. Пример результата работы программы"))
    story.append(sp(4))
    story.append(body(
        "Ниже приведён сокращённый фрагмент консольного вывода "
        "с ключевыми моментами работы программы для системы размерности n = 3."
    ))
    story.append(sp(4))
    story.append(code_block([
        "=== Test n = 3 ===",
        "Матрица системы [A|b]:",
        "   8.31  -2.14   5.67 |  12.50",
        "  -1.08   7.92   0.33 |   4.76",
        "   3.55   1.22  -6.44 |  -8.19",
        "",
        "Определитель: -312.854",
        "Решение x: (1.234, 0.567, 0.891)",
        "Норма невязки ||Ax - b|| = 2.44e-15",
    ]))
    story.append(sp(8))

    # ════════════════════════════════════════════════════════════
    # 9. ВЫВОД
    # ════════════════════════════════════════════════════════════
    story.append(heading("9. Вывод"))
    story.append(sp(4))
    story.append(body(
        "В ходе выполнения лабораторной работы был реализован метод Гаусса "
        "с частичным выбором ведущего элемента для решения систем линейных уравнений "
        "произвольной размерности."
    ))
    story.append(sp(6))
    story.append(body("Разработанная программа:"))
    story.append(sp(3))
    story.append(bullet(
        "Корректно создаёт двумерный массив (матрицу) и одномерный массив (вектор) "
        "в динамической памяти с помощью функции create."
    ))
    story.append(bullet(
        "Решает систему и вычисляет определитель матрицы с помощью функции gauss, "
        "реализующей численно устойчивый алгоритм с выбором ведущего элемента."
    ))
    story.append(bullet(
        "Обеспечивает высокую точность решения для систем размерности до n = 500 "
        "при норме невязки ~10\u207b\u2078."
    ))
    story.append(bullet(
        "Предотвращает утечки памяти за счёт явного освобождения всех "
        "динамически выделенных массивов."
    ))
    story.append(sp(8))
    story.append(body(
        "Стратегия частичного выбора ведущего элемента подтвердила свою эффективность: "
        "точность решения остаётся высокой даже для систем большой размерности, "
        "а накопление ошибок округления остаётся в допустимых пределах для задач "
        "инженерных и научных вычислений."
    ))

    doc.build(story)
    print(f"PDF сгенерирован: {output_path}")


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "/workspace/отчёт_готовый.pdf"
    build(out)
