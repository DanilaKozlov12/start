from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

# Color palette
COLOR_DARK_BLUE = RGBColor(0x1A, 0x37, 0x6C)   # deep government blue
COLOR_MED_BLUE  = RGBColor(0x2E, 0x5F, 0xA3)   # medium blue
COLOR_LIGHT_BLUE= RGBColor(0xD6, 0xE4, 0xF7)   # light blue background
COLOR_GOLD      = RGBColor(0xC9, 0xA2, 0x27)   # gold accent
COLOR_WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_DARK_TEXT = RGBColor(0x1A, 0x1A, 0x2E)

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

BLANK_LAYOUT = prs.slide_layouts[6]  # completely blank


# ── helpers ────────────────────────────────────────────────────────────────

def add_rect(slide, left, top, width, height, fill_color, line_color=None):
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
    else:
        shape.line.fill.background()
    return shape


def add_textbox(slide, text, left, top, width, height,
                font_size=18, bold=False, color=COLOR_DARK_TEXT,
                align=PP_ALIGN.LEFT, font_name="Calibri", italic=False,
                wrap=True):
    txBox = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = font_name
    return txBox


def add_bullet_box(slide, items, left, top, width, height,
                   font_size=16, color=COLOR_DARK_TEXT,
                   bullet_char="◆", font_name="Calibri", line_spacing=1.15):
    from pptx.util import Pt
    from pptx.oxml.ns import qn
    from lxml import etree

    txBox = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True

    for i, item in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.space_before = Pt(3)
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = f"{bullet_char}  {item}"
        run.font.size = Pt(font_size)
        run.font.color.rgb = color
        run.font.name = font_name
    return txBox


def slide_background(slide, top_height=1.5):
    """Dark-blue header bar + light blue body."""
    add_rect(slide, 0, 0, 13.33, top_height, COLOR_DARK_BLUE)
    add_rect(slide, 0, top_height, 13.33, 7.5 - top_height, COLOR_LIGHT_BLUE)


def gold_accent_line(slide, top_y):
    add_rect(slide, 0, top_y, 13.33, 0.07, COLOR_GOLD)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 1 – Title
# ══════════════════════════════════════════════════════════════════════════
slide1 = prs.slides.add_slide(BLANK_LAYOUT)

# Full dark-blue background
add_rect(slide1, 0, 0, 13.33, 7.5, COLOR_DARK_BLUE)
# Gold decorative lines
add_rect(slide1, 0, 2.55, 13.33, 0.07, COLOR_GOLD)
add_rect(slide1, 0, 5.3,  13.33, 0.07, COLOR_GOLD)

add_textbox(slide1,
    "Конституционные гарантии\nместного самоуправления",
    0.7, 0.6, 12.0, 1.9,
    font_size=36, bold=True, color=COLOR_WHITE,
    align=PP_ALIGN.CENTER, font_name="Calibri")

add_textbox(slide1,
    "Правовая основа местного самоуправления\nв Российской Федерации",
    0.7, 2.75, 12.0, 1.5,
    font_size=24, bold=False, color=COLOR_GOLD,
    align=PP_ALIGN.CENTER, font_name="Calibri")

add_textbox(slide1,
    "Конституционное право России",
    0.7, 5.5, 12.0, 0.7,
    font_size=16, bold=False, color=COLOR_LIGHT_BLUE,
    align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 2 – Понятие МСУ
# ══════════════════════════════════════════════════════════════════════════
slide2 = prs.slides.add_slide(BLANK_LAYOUT)
slide_background(slide2, 1.4)
gold_accent_line(slide2, 1.4)

add_textbox(slide2, "Местное самоуправление: понятие и сущность",
    0.4, 0.2, 12.5, 1.1,
    font_size=26, bold=True, color=COLOR_WHITE,
    align=PP_ALIGN.LEFT)

add_textbox(slide2,
    "«Местное самоуправление в Российской Федерации — форма осуществления народом своей власти, "
    "обеспечивающая самостоятельное и под свою ответственность решение населением непосредственно "
    "и (или) через органы МСУ вопросов местного значения»\n\n"
    "— Федеральный закон № 131-ФЗ, ст. 1",
    0.5, 1.6, 12.3, 1.9,
    font_size=15, bold=False, color=COLOR_DARK_BLUE,
    italic=True)

add_bullet_box(slide2, [
    "МСУ — конституционно закреплённая форма народовластия",
    "Самостоятельность в решении вопросов местного значения",
    "Отделено от системы государственной власти (ст. 12 Конституции РФ)",
    "Реализуется населением непосредственно и через выборные органы",
], 0.5, 3.65, 12.3, 3.3,
    font_size=16, color=COLOR_DARK_TEXT)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 3 – Статья 12 Конституции
# ══════════════════════════════════════════════════════════════════════════
slide3 = prs.slides.add_slide(BLANK_LAYOUT)
slide_background(slide3, 1.4)
gold_accent_line(slide3, 1.4)

add_textbox(slide3, "Статья 12 Конституции РФ — ключевая норма",
    0.4, 0.2, 12.5, 1.1,
    font_size=26, bold=True, color=COLOR_WHITE,
    align=PP_ALIGN.LEFT)

# Big quote box
q_box = add_rect(slide3, 0.5, 1.6, 12.3, 2.0, COLOR_MED_BLUE)
add_textbox(slide3,
    "«В Российской Федерации признаётся и гарантируется местное самоуправление. "
    "Местное самоуправление в пределах своих полномочий самостоятельно. "
    "Органы местного самоуправления не входят в систему органов государственной власти.»",
    0.7, 1.65, 11.9, 1.9,
    font_size=17, bold=False, color=COLOR_WHITE,
    align=PP_ALIGN.CENTER, italic=True)

add_bullet_box(slide3, [
    "Признание МСУ — Россия обязана создавать условия для его существования",
    "Гарантирование — государство обеспечивает реализацию права граждан на МСУ",
    "Самостоятельность — невмешательство государства в вопросы местного значения",
    "Организационная обособленность органов МСУ от органов госвласти",
], 0.5, 3.75, 12.3, 3.3,
    font_size=16, color=COLOR_DARK_TEXT)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 4 – Глава 8 Конституции (обзор)
# ══════════════════════════════════════════════════════════════════════════
slide4 = prs.slides.add_slide(BLANK_LAYOUT)
slide_background(slide4, 1.4)
gold_accent_line(slide4, 1.4)

add_textbox(slide4, "Глава 8 Конституции РФ «Местное самоуправление»",
    0.4, 0.2, 12.5, 1.1,
    font_size=26, bold=True, color=COLOR_WHITE)

articles = [
    ("Ст. 130", "Право граждан на осуществление МСУ,\nформы прямого волеизъявления"),
    ("Ст. 131", "Территориальные основы МСУ,\nучёт мнения населения при изменении границ"),
    ("Ст. 132", "Полномочия органов МСУ:\nбюджет, налоги, ЖКХ, охрана порядка"),
    ("Ст. 133", "Гарантии МСУ: судебная защита,\nкомпенсация расходов, запрет ограничений"),
]

box_w = 2.85
for i, (num, desc) in enumerate(articles):
    x = 0.4 + i * 3.15
    add_rect(slide4, x, 1.6, box_w, 1.0, COLOR_DARK_BLUE)
    add_textbox(slide4, num, x, 1.65, box_w, 0.75,
        font_size=22, bold=True, color=COLOR_GOLD,
        align=PP_ALIGN.CENTER)
    add_rect(slide4, x, 2.6, box_w, 2.0, COLOR_WHITE)
    add_textbox(slide4, desc, x + 0.1, 2.65, box_w - 0.2, 1.9,
        font_size=14, color=COLOR_DARK_TEXT,
        align=PP_ALIGN.CENTER)

add_textbox(slide4,
    "Глава 8 образует конституционный каркас системы местного самоуправления и "
    "непосредственно определяет объём прав граждан и гарантий их защиты",
    0.5, 4.8, 12.3, 1.0,
    font_size=15, italic=True, color=COLOR_MED_BLUE,
    align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 5 – Статьи 130-131
# ══════════════════════════════════════════════════════════════════════════
slide5 = prs.slides.add_slide(BLANK_LAYOUT)
slide_background(slide5, 1.4)
gold_accent_line(slide5, 1.4)

add_textbox(slide5, "Статьи 130–131: право и территориальные основы МСУ",
    0.4, 0.2, 12.5, 1.1,
    font_size=26, bold=True, color=COLOR_WHITE)

# Left column – ст.130
add_rect(slide5, 0.4, 1.6, 5.9, 5.6, COLOR_WHITE)
add_rect(slide5, 0.4, 1.6, 5.9, 0.55, COLOR_MED_BLUE)
add_textbox(slide5, "Статья 130", 0.5, 1.65, 5.7, 0.45,
    font_size=18, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
add_bullet_box(slide5, [
    "Население решает вопросы местного значения самостоятельно",
    "Формы: референдум, выборы, иное прямое волеизъявление",
    "Реализация через выборные и иные органы МСУ",
    "Право на владение, пользование и распоряжение муниципальной собственностью",
], 0.55, 2.25, 5.6, 4.8, font_size=14, color=COLOR_DARK_TEXT)

# Right column – ст.131
add_rect(slide5, 7.0, 1.6, 5.9, 5.6, COLOR_WHITE)
add_rect(slide5, 7.0, 1.6, 5.9, 0.55, COLOR_MED_BLUE)
add_textbox(slide5, "Статья 131", 7.1, 1.65, 5.7, 0.45,
    font_size=18, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
add_bullet_box(slide5, [
    "МСУ осуществляется в городских, сельских поселениях и на иных территориях",
    "Структура органов МСУ определяется населением самостоятельно",
    "Изменение границ — только с учётом мнения населения",
    "Гарантия территориальной стабильности муниципальных образований",
], 7.15, 2.25, 5.6, 4.8, font_size=14, color=COLOR_DARK_TEXT)

add_rect(slide5, 6.27, 1.6, 0.46, 5.6, COLOR_GOLD)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 6 – Статья 132: полномочия
# ══════════════════════════════════════════════════════════════════════════
slide6 = prs.slides.add_slide(BLANK_LAYOUT)
slide_background(slide6, 1.4)
gold_accent_line(slide6, 1.4)

add_textbox(slide6, "Статья 132 — Полномочия органов местного самоуправления",
    0.4, 0.2, 12.5, 1.1,
    font_size=26, bold=True, color=COLOR_WHITE)

powers = [
    ("Бюджет и финансы", "Формирование, утверждение и исполнение местного бюджета;\nустановление местных налогов и сборов"),
    ("Управление имуществом", "Владение, пользование и распоряжение муниципальной собственностью"),
    ("ЖКХ и благоустройство", "Обеспечение коммунальных услуг, дорог, транспорта,\nблагоустройства территории"),
    ("Охрана общественного порядка", "Организация охраны общественного порядка (муниципальная полиция)"),
    ("Делегированные полномочия", "Органы МСУ могут наделяться отдельными\nгосударственными полномочиями с передачей финансирования"),
]

for i, (title, desc) in enumerate(powers):
    row = i // 3
    col = i % 3
    if i >= 3:
        col = i - 3
        x = 0.4 + col * 4.3 + 1.35
    else:
        x = 0.4 + col * 4.3
    y = 1.65 + row * 2.65

    add_rect(slide6, x, y, 4.0, 0.5, COLOR_DARK_BLUE)
    add_textbox(slide6, title, x + 0.1, y + 0.05, 3.8, 0.42,
        font_size=14, bold=True, color=COLOR_GOLD, align=PP_ALIGN.CENTER)
    add_rect(slide6, x, y + 0.5, 4.0, 1.9, COLOR_WHITE)
    add_textbox(slide6, desc, x + 0.1, y + 0.55, 3.8, 1.8,
        font_size=13, color=COLOR_DARK_TEXT, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 7 – Статья 133: гарантии
# ══════════════════════════════════════════════════════════════════════════
slide7 = prs.slides.add_slide(BLANK_LAYOUT)
slide_background(slide7, 1.4)
gold_accent_line(slide7, 1.4)

add_textbox(slide7, "Статья 133 — Конституционные гарантии МСУ",
    0.4, 0.2, 12.5, 1.1,
    font_size=26, bold=True, color=COLOR_WHITE)

guarantees = [
    ("⚖️ Судебная защита",
     "Право органов МСУ и граждан на судебную защиту права на МСУ от нарушений любыми субъектами"),
    ("💰 Компенсация расходов",
     "Государство обязано компенсировать дополнительные расходы при наделении органов МСУ госполномочиями"),
    ("🚫 Запрет ограничений",
     "Запрет на ограничение прав МСУ, установленных Конституцией и федеральными законами"),
]

for i, (title, desc) in enumerate(guarantees):
    y = 1.6 + i * 1.8
    add_rect(slide7, 0.4, y, 12.4, 1.55, COLOR_WHITE)
    add_rect(slide7, 0.4, y, 0.25, 1.55, COLOR_GOLD)
    add_textbox(slide7, title, 0.8, y + 0.1, 3.5, 0.55,
        font_size=17, bold=True, color=COLOR_DARK_BLUE)
    add_textbox(slide7, desc, 0.8, y + 0.65, 11.8, 0.8,
        font_size=15, color=COLOR_DARK_TEXT)

add_textbox(slide7,
    "Данные гарантии носят конституционный характер и не могут быть отменены ни федеральным, ни региональным законом",
    0.5, 7.0, 12.3, 0.45,
    font_size=13, italic=True, color=COLOR_MED_BLUE, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 8 – 131-ФЗ
# ══════════════════════════════════════════════════════════════════════════
slide8 = prs.slides.add_slide(BLANK_LAYOUT)
slide_background(slide8, 1.4)
gold_accent_line(slide8, 1.4)

add_textbox(slide8, "Федеральный закон № 131-ФЗ от 06.10.2003",
    0.4, 0.2, 12.5, 1.1,
    font_size=26, bold=True, color=COLOR_WHITE)

add_textbox(slide8,
    "«Об общих принципах организации местного самоуправления в Российской Федерации»",
    0.5, 1.5, 12.3, 0.7,
    font_size=18, bold=True, color=COLOR_DARK_BLUE, align=PP_ALIGN.CENTER, italic=True)

add_bullet_box(slide8, [
    "Базовый федеральный закон, конкретизирующий конституционные положения о МСУ",
    "Устанавливает виды муниципальных образований: городские и сельские поселения, муниципальные районы, городские округа, внутригородские территории",
    "Закрепляет перечень вопросов местного значения для каждого типа муниципального образования",
    "Регулирует систему органов МСУ: представительный орган, глава, администрация, контрольный орган",
    "Определяет экономическую основу МСУ: муниципальное имущество, местный бюджет, налоги",
    "Устанавливает формы непосредственного участия граждан: местный референдум, муниципальные выборы, сход граждан, правотворческая инициатива, публичные слушания",
], 0.5, 2.3, 12.3, 4.9, font_size=15, color=COLOR_DARK_TEXT)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 9 – Европейская Хартия
# ══════════════════════════════════════════════════════════════════════════
slide9 = prs.slides.add_slide(BLANK_LAYOUT)
slide_background(slide9, 1.4)
gold_accent_line(slide9, 1.4)

add_textbox(slide9, "Европейская хартия местного самоуправления 1985 г.",
    0.4, 0.2, 12.5, 1.1,
    font_size=26, bold=True, color=COLOR_WHITE)

add_textbox(slide9,
    "Ратифицирована Россией в 1998 году (Федеральный закон № 55-ФЗ) — является частью правовой системы РФ",
    0.5, 1.5, 12.3, 0.65,
    font_size=15, bold=True, color=COLOR_DARK_BLUE, italic=True, align=PP_ALIGN.CENTER)

# Two columns
add_rect(slide9, 0.4, 2.3, 5.9, 4.8, COLOR_WHITE)
add_rect(slide9, 0.4, 2.3, 5.9, 0.55, COLOR_MED_BLUE)
add_textbox(slide9, "Ключевые принципы Хартии", 0.5, 2.35, 5.7, 0.45,
    font_size=16, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
add_bullet_box(slide9, [
    "Право и реальная способность органов МСУ самостоятельно регулировать местные дела",
    "Принцип субсидиарности: вопросы решаются на максимально близком к гражданину уровне",
    "Финансовая самостоятельность — собственные налоги и сборы",
    "Административный надзор только в пределах, установленных законом",
    "Судебная защита автономии МСУ",
], 0.55, 2.95, 5.65, 4.0, font_size=13, color=COLOR_DARK_TEXT)

add_rect(slide9, 7.0, 2.3, 5.9, 4.8, COLOR_WHITE)
add_rect(slide9, 7.0, 2.3, 5.9, 0.55, COLOR_MED_BLUE)
add_textbox(slide9, "Значение для России", 7.1, 2.35, 5.7, 0.45,
    font_size=16, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
add_bullet_box(slide9, [
    "Источник толкования конституционных норм о МСУ",
    "Стандарты организации МСУ имплементированы в 131-ФЗ",
    "Приоритет перед нормами национального законодательства (ч. 4 ст. 15 Конституции)",
    "Основа для обращений в международные органы",
    "После 2022 года — правовой статус нормы пересматривается",
], 7.15, 2.95, 5.65, 4.0, font_size=13, color=COLOR_DARK_TEXT)

add_rect(slide9, 6.27, 2.3, 0.46, 4.8, COLOR_GOLD)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 10 – Иные источники права
# ══════════════════════════════════════════════════════════════════════════
slide10 = prs.slides.add_slide(BLANK_LAYOUT)
slide_background(slide10, 1.4)
gold_accent_line(slide10, 1.4)

add_textbox(slide10, "Иные источники правового регулирования МСУ",
    0.4, 0.2, 12.5, 1.1,
    font_size=26, bold=True, color=COLOR_WHITE)

levels = [
    ("Федеральный уровень",
     ["ФЗ № 67-ФЗ — об основных гарантиях избирательных прав",
      "ФЗ № 25-ФЗ — о муниципальной службе",
      "ФЗ № 154-ФЗ — о финансовых основах МСУ (утр. силу, заменён бюджетным законодательством)",
      "Бюджетный и Налоговый кодексы РФ — финансовые основы МСУ",
      "Земельный кодекс РФ — управление земельными ресурсами"],
     COLOR_DARK_BLUE),
    ("Региональный уровень",
     ["Конституции и Уставы субъектов РФ",
      "Законы субъектов РФ об организации МСУ",
      "Законы об установлении границ муниципальных образований",
      "Законы о наделении органов МСУ отдельными государственными полномочиями"],
     COLOR_MED_BLUE),
    ("Муниципальный уровень",
     ["Устав муниципального образования — основной акт МСУ",
      "Решения представительного органа",
      "Постановления и распоряжения главы и администрации"],
     RGBColor(0x4A, 0x90, 0xD9)),
]

col_w = 4.0
for i, (title, items, color) in enumerate(levels):
    x = 0.4 + i * 4.3
    add_rect(slide10, x, 1.6, col_w, 0.6, color)
    add_textbox(slide10, title, x + 0.1, 1.63, col_w - 0.2, 0.55,
        font_size=14, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
    add_rect(slide10, x, 2.2, col_w, 5.0, COLOR_WHITE)
    add_bullet_box(slide10, items, x + 0.15, 2.3, col_w - 0.25, 4.8,
        font_size=13, color=COLOR_DARK_TEXT, bullet_char="•")


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 11 – Конституционный суд о МСУ
# ══════════════════════════════════════════════════════════════════════════
slide11 = prs.slides.add_slide(BLANK_LAYOUT)
slide_background(slide11, 1.4)
gold_accent_line(slide11, 1.4)

add_textbox(slide11, "Роль Конституционного Суда РФ в защите МСУ",
    0.4, 0.2, 12.5, 1.1,
    font_size=26, bold=True, color=COLOR_WHITE)

add_textbox(slide11,
    "Конституционный Суд РФ является ключевым органом защиты конституционных гарантий МСУ",
    0.5, 1.55, 12.3, 0.6,
    font_size=16, bold=True, color=COLOR_DARK_BLUE, italic=True, align=PP_ALIGN.CENTER)

decisions = [
    ("Постановление КС РФ\n№ 9-П (1997)",
     "Закрепил принцип организационной самостоятельности МСУ и недопустимость произвольного вмешательства государства"),
    ("Постановление КС РФ\n№ 15-П (2000)",
     "О праве населения на МСУ при изменении границ — без согласия населения изменения незаконны"),
    ("Постановление КС РФ\n№ 13-П (2002)",
     "Подтвердил, что органы МСУ не могут быть включены в систему государственной власти"),
    ("Постановление КС РФ\n№ 20-П (2015)",
     "О допустимости введения сити-менеджера при сохранении выборного начала МСУ"),
]

for i, (title, desc) in enumerate(decisions):
    row = i // 2
    col = i % 2
    x = 0.4 + col * 6.5
    y = 2.3 + row * 2.5
    add_rect(slide11, x, y, 6.1, 2.2, COLOR_WHITE)
    add_rect(slide11, x, y, 6.1, 0.08, COLOR_GOLD)
    add_textbox(slide11, title, x + 0.15, y + 0.15, 5.8, 0.75,
        font_size=14, bold=True, color=COLOR_DARK_BLUE)
    add_textbox(slide11, desc, x + 0.15, y + 0.9, 5.8, 1.2,
        font_size=13, color=COLOR_DARK_TEXT)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 12 – Система гарантий МСУ
# ══════════════════════════════════════════════════════════════════════════
slide12 = prs.slides.add_slide(BLANK_LAYOUT)
slide_background(slide12, 1.4)
gold_accent_line(slide12, 1.4)

add_textbox(slide12, "Система конституционных гарантий МСУ",
    0.4, 0.2, 12.5, 1.1,
    font_size=26, bold=True, color=COLOR_WHITE)

guarantee_groups = [
    ("Организационные\nгарантии",
     ["Самостоятельное определение структуры органов МСУ", "Организационная обособленность от госвласти", "Запрет произвольного роспуска органов МСУ"],
     COLOR_DARK_BLUE),
    ("Финансово-экономические\nгарантии",
     ["Право на местный бюджет", "Собственные налоговые доходы", "Муниципальная собственность", "Компенсация делегированных полномочий"],
     COLOR_MED_BLUE),
    ("Юридические\nгарантии",
     ["Судебная защита прав МСУ", "Право на обращение в КС РФ", "Ответственность за нарушение прав МСУ", "Запрет ограничения прав МСУ"],
     RGBColor(0x4A, 0x90, 0xD9)),
    ("Политические\nгарантии",
     ["Выборность органов МСУ", "Право на местный референдум", "Сход граждан", "Публичные слушания"],
     RGBColor(0x5B, 0xA5, 0x7D)),
]

col_w = 3.0
for i, (title, items, color) in enumerate(guarantee_groups):
    x = 0.27 + i * 3.2
    add_rect(slide12, x, 1.6, col_w, 0.75, color)
    add_textbox(slide12, title, x + 0.1, 1.63, col_w - 0.2, 0.7,
        font_size=13, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
    add_rect(slide12, x, 2.35, col_w, 4.85, COLOR_WHITE)
    add_bullet_box(slide12, items, x + 0.12, 2.45, col_w - 0.22, 4.65,
        font_size=13, color=COLOR_DARK_TEXT, bullet_char="•")


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 13 – Заключение
# ══════════════════════════════════════════════════════════════════════════
slide13 = prs.slides.add_slide(BLANK_LAYOUT)
add_rect(slide13, 0, 0, 13.33, 7.5, COLOR_DARK_BLUE)
add_rect(slide13, 0, 2.1, 13.33, 0.07, COLOR_GOLD)
add_rect(slide13, 0, 5.55, 13.33, 0.07, COLOR_GOLD)

add_textbox(slide13, "Выводы",
    0.7, 0.25, 12.0, 0.9,
    font_size=34, bold=True, color=COLOR_WHITE,
    align=PP_ALIGN.CENTER)

add_bullet_box(slide13, [
    "МСУ является конституционной формой народовластия (ст. 3, 12, гл. 8 Конституции РФ)",
    "Конституция РФ закрепляет как само право на МСУ, так и систему его гарантий",
    "ФЗ № 131-ФЗ детализирует конституционные нормы и формирует законодательную базу",
    "Европейская хартия МСУ служила международным стандартом до изменения правового статуса",
    "КС РФ — главный страж конституционных гарантий МСУ в России",
], 0.8, 2.3, 11.7, 3.0,
    font_size=17, color=COLOR_WHITE, bullet_char="✓")

add_textbox(slide13,
    "«Местное самоуправление — это демократия, которую гражданин ощущает каждый день»",
    0.7, 5.75, 12.0, 0.9,
    font_size=16, italic=True, color=COLOR_GOLD, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════
# SAVE
# ══════════════════════════════════════════════════════════════════════════
out_path = "/workspace/Конституционные_гарантии_МСУ.pptx"
prs.save(out_path)
print(f"Saved: {out_path}")
