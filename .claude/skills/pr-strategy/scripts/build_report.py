# -*- coding: utf-8 -*-
"""
[LEGACY] Программный генератор PR-отчёта с нуля (python-pptx, канва 13.33×7.5).

⚠️ КАНОНИЧЕСКИЙ ШАБЛОН — assets/report-template.pptx (утверждён агентством, канва
10×5.625). Будущие презентации собираем на его основе (копия + замена плейсхолдеров),
а не этим скриптом. Скрипт оставлен для справки/возможной автоматизации; при
расхождении приоритет у мастер-шаблона.

Генератор ежемесячного PR-отчёта в фирменном стиле GORD AGENCY (16:9, pptx).
Реализация стандарта из references/gord-visual-identity.md и
references/monthly-report-template.md. Основа — разбор эталонов Domingo Dacha / Честная Рыба.

СТРУКТУРА (12 слайдов): обложка → результаты (дашборд) → выполненные PR-задачи →
AVE/OTS ×4 (СМИ, ТГ, INST, INST/ТТ/Max) → пресс-релиз → публикации ×3 (СМИ, ТГ, INST)
→ финал. Слайды с данными за месяц идут пустыми плейсхолдерами — заполняются фактурой.

ЗАВИСИМОСТИ:
  - Python 3 + python-pptx  (pip install python-pptx)
  - Корпоративный шрифт Geologica установлен в системе (иначе PowerPoint подставит другой).
    Файл: «Шрифты (корпоративные)/Geologica-VariableFont_CRSV,SHRP,slnt,wght.ttf».

НАСТРОЙКА: правь блок «Данные проекта» ниже (BRAND, PERIOD, YEAR, SINCE, OUT).
ЗАПУСК:   python build_report.py   → создаёт pptx по пути OUT.

Стандарт оформления (кратко): красный #D72410, чёрные тёмные страницы, бежевый дашборд
#CDB996; заголовки Geologica 28pt UPPER без жирного, белые; строчная «x» на обложке;
вертикальный вотермарк «GORD AGENCY» впритык к правому краю до верх/низ границ; контент
не заходит под вотермарк; год — в нижнем-левом углу (отступ 1.5 см), кроме дашборда;
нумерации слайдов нет.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---------- Бренд-константы ----------
RED    = RGBColor(0xD7, 0x24, 0x10)   # Signal Clay
BLACK  = RGBColor(0x00, 0x00, 0x00)   # фон тёмных страниц
RICH   = RGBColor(0x10, 0x0E, 0x09)   # Rich Black (текст)
BEIGE  = RGBColor(0xCD, 0xB9, 0x96)   # Warm Beige (дашборд)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
GREY_H = RGBColor(0xD9, 0xD9, 0xD9)   # шапка таблицы
GREYTX = RGBColor(0x8A, 0x8A, 0x8A)   # приглушённый текст/плейсхолдер
BLUE   = RGBColor(0x2F, 0x6F, 0xE0)   # ссылки площадок
WM_DK  = RGBColor(0x1A, 0x1A, 0x1A)   # тёмный вотермарк на фото

FONT_H = "Geologica"             # заголовки, числа (в т.ч. bold)
FONT_B = "Geologica"             # основной текст
FONT_L = "Geologica ExtraLight"  # лёгкие подписи (подзаголовки, вторые строки KPI)

# ---------- Данные проекта (правь под клиента/период) ----------
BRAND   = "A-FERMA"          # название бренда (UPPER), пойдёт в «GORD AGENCY x BRAND»
PERIOD  = "ИЮНЬ 2026"        # отчётный период (подзаголовок обложки)
YEAR    = "2026"             # актуальный год (нижний-левый колонтитул)
SINCE   = "SINCE 2022"       # год старта сотрудничества (нижний-правый угол обложки)
OUT     = "a_ferma_report.pptx"   # путь сохранения pptx

EMU_W, EMU_H = Inches(13.333), Inches(7.5)

prs = Presentation()
prs.slide_width  = EMU_W
prs.slide_height = EMU_H
BLANK = prs.slide_layouts[6]

# ---------- helpers ----------
def bg(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color

def add_text(slide, x, y, w, h, text, size, color, bold=False, italic=False,
             align=PP_ALIGN.LEFT, font=FONT_B, anchor=MSO_ANCHOR.TOP,
             spacing=None, upper=False, line_spacing=None):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    lines = text.split("\n")
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if line_spacing: p.line_spacing = line_spacing
        r = p.add_run()
        r.text = ln.upper() if upper else ln
        f = r.font
        f.size = Pt(size); f.bold = bold; f.italic = italic
        f.name = font; f.color.rgb = color
        if spacing is not None:
            rPr = r._r.get_or_add_rPr()
            rPr.set('spc', str(int(spacing)))
    return tb

def rect(slide, x, y, w, h, fill=None, line=None, line_w=Pt(1), dash=False):
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    sp.shadow.inherit = False
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line; sp.line.width = line_w
        if dash:
            ln = sp.line._get_or_add_ln()
            d = ln.makeelement(qn('a:prstDash'), {'val': 'dash'})
            ln.append(d)
    return sp

def corners(slide, tl="", tr="", bl="", br="", color=WHITE, br_inset=Inches(0.4), bl_inset=Inches(0.4)):
    # bl_inset — левый отступ для нижнего-левого лейбла; br_inset — правый отступ для нижнего-правого
    m = Inches(0.4)
    if tl: add_text(slide, m, m, Inches(4), Inches(0.3), tl, 9, color, font=FONT_B, upper=True)
    if tr: add_text(slide, EMU_W-Inches(4.4), m, Inches(4), Inches(0.3), tr, 9, color, align=PP_ALIGN.RIGHT, font=FONT_B, upper=True)
    if bl: add_text(slide, bl_inset, EMU_H-Inches(0.55), Inches(4), Inches(0.3), bl, 9, color, font=FONT_B, upper=True)
    if br: add_text(slide, EMU_W-Inches(4)-br_inset, EMU_H-Inches(0.55), Inches(4), Inches(0.3), br, 9, color, align=PP_ALIGN.RIGHT, font=FONT_B, upper=True)

def watermark(slide, color=RED):
    """Вертикальная надпись GORD AGENCY впритык к правому краю (поворот 270°),
    упирается в верхнюю и нижнюю границы слайда.
    Бокс W×H после поворота даёт визуальный H×W с центром в центре бокса:
    W = вертикальная длина (= высота слайда), H = толщина строки."""
    W, H = EMU_H, Inches(1.18)             # W = полная высота слайда; H ≈ высоте букв, чтобы не было зазора
    cx = EMU_W - Emu(int(H/2)) + Inches(0.02) + Cm(0.5)   # ещё +0.5 см вправо, по границе листа
    cy = int(EMU_H/2)
    x, y = int(cx - W/2), int(cy - H/2)
    tb = slide.shapes.add_textbox(x, y, W, H)
    tb.rotation = 270
    tf = tb.text_frame; tf.word_wrap = False; tf.auto_size = None
    tf.margin_left=0; tf.margin_right=0; tf.margin_top=0; tf.margin_bottom=0
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = "GORD AGENCY"
    r.font.size = Pt(72); r.font.bold = True; r.font.name = FONT_H; r.font.color.rgb = color

def note(slide, x, y, w, h, text, color=GREYTX):
    """Пунктирная рамка-плейсхолдер с подписью по центру."""
    rect(slide, x, y, w, h, fill=None, line=color, line_w=Pt(1.25), dash=True)
    add_text(slide, x+Inches(0.2), y, w-Inches(0.4), h, text, 15, color,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font=FONT_B, italic=True)

def title(slide, text, color=WHITE, sub=None, sub_color=None):
    # размер 28, поднято на 1 см (0.85" - 0.394")
    add_text(slide, Inches(0.6), Inches(0.456), Inches(10.6), Inches(0.9), text, 28,
             color, bold=False, font=FONT_H, upper=True, spacing=30)
    if sub:
        add_text(slide, Inches(0.62), Inches(1.206), Inches(10.6), Inches(0.5), sub,
                 15, sub_color or color, font=FONT_L, upper=True, spacing=20)

# =================== СЛАЙД 1 — ОБЛОЖКА ===================
s = prs.slides.add_slide(BLANK)
bg(s, RICH)
# место под фон-фото объекта
rect(s, 0, 0, EMU_W, EMU_H, fill=RICH)
# строчная «x» между названиями; заголовок жирный (стиль эталона), Geologica
add_text(s, Inches(0.6), Inches(3.0), Inches(12.1), Inches(1.2),
         f"GORD AGENCY x {BRAND}", 52, RED, bold=True, font=FONT_H,
         align=PP_ALIGN.CENTER, upper=False)
add_text(s, Inches(0.6), Inches(4.2), Inches(12.1), Inches(0.5),
         PERIOD, 18, WHITE, font=FONT_L, align=PP_ALIGN.CENTER, upper=True)
corners(s, tr=YEAR, br=SINCE, color=WHITE)
# служебная подпись про фон (по центру, чтобы не пересекаться с годом слева)
add_text(s, Inches(0), EMU_H-Inches(0.5), EMU_W, Inches(0.3),
         f"[ фон: фото объекта {BRAND} с тёмным оверлеем 40–60% ]", 8, GREYTX,
         italic=True, align=PP_ALIGN.CENTER)

# =================== СЛАЙД 2 — РЕЗУЛЬТАТЫ (дашборд) ===================
s = prs.slides.add_slide(BLANK)
bg(s, BEIGE)
add_text(s, Inches(0.6), Inches(0.306), Inches(6), Inches(0.8), "РЕЗУЛЬТАТЫ", 28,
         WHITE, bold=False, font=FONT_H, upper=True, spacing=30)
corners(s, tr="GORDPR.RU", bl="GORD AGENCY", color=WHITE)

def kpi(y, big, lab1, lab2):
    rect(s, Inches(0.62), y, Pt(3), Inches(1.15), fill=WHITE)   # вертикальная линейка
    add_text(s, Inches(0.85), y-Inches(0.05), Inches(5.5), Inches(0.7), big, 36,
             WHITE, bold=True, font=FONT_H)
    add_text(s, Inches(0.87), y+Inches(0.58), Inches(5.5), Inches(0.32), lab1, 16,
             WHITE, font=FONT_H, upper=True)          # первая строка — Geologica, UPPER
    add_text(s, Inches(0.87), y+Inches(0.9), Inches(5.5), Inches(0.32), lab2, 16,
             WHITE, font=FONT_L)                        # вторая строка — ExtraLight, sentence

kpi(Inches(1.85), "—",         "ОБЩИЙ ОХВАТ", "за месяц")
kpi(Inches(3.55), "—",         "10% ОТ ОХВАТА", "в СМИ за месяц")
kpi(Inches(5.25), "— руб.",    "ВЫГОДА", "PR-value")

# правый блок — МЕСТО ПОД ФОТО (тёмная подложка, чтобы белые KPI читались, + рамка-плейсхолдер)
bx, by, bw, bh = Inches(7.0), Inches(1.55), Inches(5.7), Inches(5.15)
rect(s, bx, by, bw, bh, fill=RICH)
rect(s, bx, by, bw, bh, fill=None, line=WHITE, line_w=Pt(1.25), dash=True)
add_text(s, bx, by+bh-Inches(0.42), bw, Inches(0.3),
         "[ МЕСТО ПОД ФОТО ]", 9, GREYTX, italic=True, align=PP_ALIGN.CENTER)
def mini(cx, cy, num, l1, l2):
    add_text(s, cx, cy, Inches(2.6), Inches(0.55), num, 30, WHITE, bold=True, font=FONT_H)
    add_text(s, cx, cy+Inches(0.52), Inches(2.6), Inches(0.3), l1, 13, WHITE,
             font=FONT_H, upper=True)
    add_text(s, cx, cy+Inches(0.82), Inches(2.6), Inches(0.5), l2, 13, WHITE, font=FONT_L)
mini(bx+Inches(0.5), by+Inches(0.45), "—", "упоминаний", "во всех медиа")
mini(bx+Inches(3.1), by+Inches(0.45), "—", "публикаций", "в СМИ")
mini(bx+Inches(0.5), by+Inches(2.6), "—", "публикаций", "в ТГ")
mini(bx+Inches(3.1), by+Inches(2.6), "—", "упоминаний", "в INST и др.")

# =================== СЛАЙД 3 — ВЫПОЛНЕННЫЕ PR-ЗАДАЧИ (пусто) ===================
s = prs.slides.add_slide(BLANK)
bg(s, BLACK)
title(s, "ВЫПОЛНЕННЫЕ PR-ЗАДАЧИ", color=WHITE, sub=f"ПРОЕКТ «{BRAND}»", sub_color=WHITE)
watermark(s, color=RED)
note(s, Inches(0.6), Inches(2.5), Inches(9.3), Inches(4.2),
     "[ список выполненных PR-задач за месяц — заполнить ]")
corners(s, bl=YEAR, color=WHITE, bl_inset=Cm(1.5))

# =================== СЛАЙДЫ AVE/OTS (пустые, по каналам) ===================
# CONTENT_R — правая граница контента, чтобы не залезать под вертикальный вотермарк
CONTENT_R = Inches(11.5)
def ave_slide(name):
    s = prs.slides.add_slide(BLANK)
    bg(s, BLACK)
    title(s, name, color=WHITE)
    watermark(s, color=RED)
    # шапка таблицы (сумма ширин = 10.9" → правый край 11.5", вне зоны вотермарка)
    cols = [("№",0.8),("Название",3.0),("Ссылка",1.8),("Формат",2.3),("AVE",1.5),("OTS",1.5)]
    hx, hy, hh = Inches(0.6), Inches(2.2), Cm(1.0)   # высота шапки таблицы = 1 см
    cx = hx
    for nm, wv in cols:
        rect(s, cx, hy, Inches(wv), hh, fill=GREY_H, line=BLACK, line_w=Pt(1))
        add_text(s, cx, hy, Inches(wv), hh, nm, 14, RICH, bold=True, font=FONT_B,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        cx += Inches(wv)
    note(s, hx, Inches(2.9), Inches(10.9), Inches(3.6),
         "[ реестр выходов за месяц: № · площадка · ссылка · формат · AVE · OTS — заполнить ]")
    corners(s, bl=YEAR, color=WHITE, bl_inset=Cm(1.5))

for nm in ["AVE/OTS В СМИ", "AVE/OTS В ТГ", "AVE/OTS В INST", "AVE/OTS В INST/ТТ/Max"]:
    ave_slide(nm)

# =================== СЛАЙД — ПРЕСС-РЕЛИЗ (без наполнения) ===================
s = prs.slides.add_slide(BLANK)
bg(s, BLACK)
title(s, "Пресс-релиз", color=WHITE)
watermark(s, color=RED)
corners(s, bl=YEAR, color=WHITE, bl_inset=Cm(1.5))   # год слева-внизу, отступ 1.5 см
# без наполнения

# =================== СЛАЙДЫ ПУБЛИКАЦИЙ (пустые, по каналам) ===================
def pub_slide(channel):
    s = prs.slides.add_slide(BLANK)
    bg(s, BLACK)
    # заголовок одним боксом: белый текст + синяя ссылка-площадка (28pt, поднят на 1 см)
    _tb = s.shapes.add_textbox(Inches(0.6), Inches(0.456), Inches(11.0), Inches(0.8))
    _tf = _tb.text_frame; _tf.word_wrap = True
    _tf.margin_left=0; _tf.margin_right=0; _tf.margin_top=0; _tf.margin_bottom=0
    _p = _tf.paragraphs[0]
    _r1 = _p.add_run(); _r1.text = f"ПУБЛИКАЦИЯ В {channel}: "
    _r1.font.size=Pt(28); _r1.font.bold=False; _r1.font.name=FONT_H; _r1.font.color.rgb=WHITE
    _r2 = _p.add_run(); _r2.text = "площадка"
    _r2.font.size=Pt(28); _r2.font.bold=False; _r2.font.name=FONT_H; _r2.font.color.rgb=BLUE
    _r2.font.underline = True
    watermark(s, color=RED)
    note(s, Inches(3.5), Inches(2.1), Inches(6.3), Inches(4.2),
         "[ скриншот публикации ]")
    add_text(s, Inches(0.6), EMU_H-Inches(1.4)+Cm(1.0), Inches(5), Inches(0.5),
             "ОХВАТ:  —", 22, WHITE, bold=True, font=FONT_H, upper=True)
    corners(s, bl=YEAR, color=WHITE, bl_inset=Cm(1.5))

for ch in ["СМИ", "ТГ", "INST"]:
    pub_slide(ch)

# =================== СЛАЙД 6 — ФИНАЛ ===================
s = prs.slides.add_slide(BLANK)
bg(s, RED)
add_text(s, Inches(0.6), Inches(3.0), Inches(12.1), Inches(0.9),
         "GORD AGENCY", 44, WHITE, bold=False, font=FONT_H, align=PP_ALIGN.CENTER, upper=True)
add_text(s, Inches(0.6), Inches(3.95), Inches(12.1), Inches(0.5),
         "БУТИКОВОЕ КОММУНИКАЦИОННОЕ АГЕНТСТВО", 14, WHITE, font=FONT_B,
         align=PP_ALIGN.CENTER, upper=True, spacing=30)
corners(s, bl=YEAR, color=WHITE, bl_inset=Cm(1.5))

prs.save(OUT)
print("SAVED", OUT, "slides:", len(prs.slides._sldIdLst))
