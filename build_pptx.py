"""Побудова презентації ЛР5 «Вебсайт» — OZK Logistics (Захарчук О.В.)
Запуск: python build_pptx.py [github-username]"""
import sys, pathlib, qrcode
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

USER = sys.argv[1] if len(sys.argv) > 1 else "your-username"
REPO = "ozk-logistics"
SITE_URL = f"https://{USER}.github.io/{REPO}/"
REPO_URL = f"https://github.com/{USER}/{REPO}"

ROOT = pathlib.Path(__file__).parent
SC = ROOT / "screens"
NAVY = RGBColor(0x0B, 0x1F, 0x3A); NAVY2 = RGBColor(0x13, 0x2D, 0x52)
ORANGE = RGBColor(0xFF, 0x7A, 0x00); WHITE = RGBColor(255, 255, 255)
TEXT = RGBColor(0x1B, 0x24, 0x33); MUTED = RGBColor(0x5B, 0x6B, 0x82)
SOFT = RGBColor(0xF3, 0xF6, 0xFA); LIGHT = RGBColor(0xC3, 0xD1, 0xE8)
GREEN = RGBColor(0x16, 0xA3, 0x4A); RED = RGBColor(0xDC, 0x26, 0x26)
HF, BF = "Montserrat", "Segoe UI"

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
BLANK = prs.slide_layouts[6]
TOTAL = 12


def rect(s, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE, line=None):
    r = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    r.fill.solid(); r.fill.fore_color.rgb = fill
    if line: r.line.color.rgb = line; r.line.width = Pt(1)
    else: r.line.fill.background()
    r.shadow.inherit = False
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE: r.adjustments[0] = 0.08
    return r


def text(s, x, y, w, h, content, size=16, color=TEXT, bold=False, font=BF, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, spacing=1.1):
    """content: str or list of str / (str, dict) runs per paragraph"""
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.05); tf.margin_top = tf.margin_bottom = Inches(0.03)
    paras = content if isinstance(content, list) else [content]
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align; para.line_spacing = spacing
        runs = p if isinstance(p, list) else [p]
        for r in runs:
            t, o = (r, {}) if isinstance(r, str) else r
            run = para.add_run(); run.text = t
            f = run.font; f.size = Pt(o.get("size", size)); f.bold = o.get("bold", bold)
            f.color.rgb = o.get("color", color); f.name = o.get("font", font)
        if isinstance(p, list) or isinstance(p, str):
            para.space_after = Pt(o.get("after", 6) if not isinstance(p, str) else 6)
    return tb


def bullets(s, x, y, w, h, items, size=15, color=TEXT, marker="●", mcolor=ORANGE):
    paras = []
    for it in items:
        if isinstance(it, tuple):
            head, body = it
            paras.append([(marker + "  ", {"color": mcolor, "size": size - 3}), (head, {"bold": True}), (" " + body, {})])
        else:
            paras.append([(marker + "  ", {"color": mcolor, "size": size - 3}), (it, {})])
    return text(s, x, y, w, h, paras, size=size, color=color, spacing=1.05)


def header(s, title, eyebrow, n):
    rect(s, 0, 0, 13.333, 7.5, WHITE)
    rect(s, 0, 0, 0.18, 7.5, ORANGE)
    text(s, 0.6, 0.35, 10, 0.4, eyebrow.upper(), size=12, color=ORANGE, bold=True, font=HF)
    text(s, 0.6, 0.65, 12, 0.9, title, size=28, color=NAVY, bold=True, font=HF)
    footer(s, n, dark=False)


def footer(s, n, dark):
    c = LIGHT if dark else MUTED
    text(s, 0.6, 7.05, 9, 0.3, "ЛР № 5 «Вебсайт» · Захарчук О. В., ФЕМП 3-16 · OZK Logistics", size=10, color=c)
    text(s, 11.9, 7.05, 0.9, 0.3, f"{n} / {TOTAL}", size=10, color=c, align=PP_ALIGN.RIGHT)


def picture(s, path, x, y, w=None, h=None, border=True):
    if border:
        # рамка "вікно браузера"
        pw = w; ph = h
        pic = s.shapes.add_picture(str(path), Inches(x), Inches(y + 0.32), Inches(w) if w else None, Inches(h) if h else None)
        pw = pic.width / 914400; ph = pic.height / 914400
        bar = rect(s, x, y, pw, 0.32, RGBColor(0xE2, 0xE8, 0xF0))
        for i, col in enumerate([RGBColor(0xFF, 0x5F, 0x57), RGBColor(0xFE, 0xBC, 0x2E), RGBColor(0x28, 0xC8, 0x40)]):
            rect(s, x + 0.15 + i * 0.22, y + 0.1, 0.13, 0.13, col, MSO_SHAPE.OVAL)
        frame = rect(s, x, y, pw, ph + 0.32, WHITE, line=RGBColor(0xCB, 0xD5, 0xE1)); frame.fill.background()
        return pic, pw, ph + 0.32
    pic = s.shapes.add_picture(str(path), Inches(x), Inches(y), Inches(w) if w else None, Inches(h) if h else None)
    return pic, pic.width / 914400, pic.height / 914400


def url_bar(s, x, y, w, url):
    text(s, x + 0.85, y + 0.04, w - 1, 0.26, url, size=10, color=MUTED)


# ================= 1. Титульний =================
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.333, 7.5, NAVY)
rect(s, 8.6, 0, 4.733, 7.5, NAVY2)
rect(s, 0.8, 1.0, 0.12, 2.3, ORANGE)
text(s, 1.15, 0.95, 7.5, 0.4, "ДЕРЖАВНИЙ ТОРГОВЕЛЬНО-ЕКОНОМІЧНИЙ УНІВЕРСИТЕТ", size=13, color=ORANGE, bold=True, font=HF)
text(s, 1.15, 1.35, 7.5, 1.0, "Лабораторна робота № 5", size=36, color=WHITE, bold=True, font=HF)
text(s, 1.15, 2.2, 7.3, 1.2, "Вебсайт для бізнес-ідеї «OZK Logistics»: аналіз аналогів та розробка сайту за допомогою ШІ", size=19, color=LIGHT)
info = [("Виконав:", "Захарчук Олександр Валентинович"), ("Факультет / курс:", "ФЕМП, 3 курс"), ("Група:", "16 (ФЕМП 3-16)"),
        ("Дисципліна:", "Інтернет-технології в бізнесі")]
for i, (k, v) in enumerate(info):
    text(s, 1.15, 4.0 + i * 0.55, 2.4, 0.5, k, size=15, color=LIGHT)
    text(s, 3.4, 4.0 + i * 0.55, 5.2, 0.5, v, size=16, color=WHITE, bold=True)
text(s, 1.15, 6.6, 6, 0.4, "Київ — 2026", size=13, color=LIGHT)
logo = rect(s, 9.55, 2.4, 2.8, 1.25, ORANGE, MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, 9.55, 2.45, 2.8, 1.15, "OZK", size=54, color=NAVY, bold=True, font=HF, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(s, 9.0, 3.8, 3.9, 0.6, "Logistics", size=30, color=WHITE, bold=True, font=HF, align=PP_ALIGN.CENTER)
text(s, 9.0, 4.5, 3.9, 0.9, "B2B-вантажоперевезення\nУкраїна ⇄ ЄС", size=15, color=LIGHT, align=PP_ALIGN.CENTER)

# ================= 2. Тематика =================
s = prs.slides.add_slide(BLANK)
header(s, "Тематика та тип майбутнього сайту", "Бізнес-ідея", 2)
rect(s, 0.6, 1.7, 6.0, 5.1, SOFT, MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, 0.9, 1.9, 5.5, 0.5, "Тематика", size=14, color=ORANGE, bold=True, font=HF)
text(s, 0.9, 2.3, 5.5, 0.9, "Логістична компанія ТОВ «ОЗК ЛОГІСТИК» (OZK Logistics)", size=18, color=NAVY, bold=True, font=HF)
bullets(s, 0.9, 3.35, 5.5, 3.4, [
    ("Послуги:", "експрес-доставка FTL/LTL по Україні, міжнародні рейси до Польщі, Німеччини, Румунії, експедирування та митний супровід, CMR-страхування і GPS-трекінг."),
    ("Цільова аудиторія (B2B):", "МСБ-виробники, e-commerce та дистриб’ютори, імпортери/експортери."),
    ("Ціннісна пропозиція:", "точність і безпека доставки за заздалегідь зафіксованим тарифом + трекінг 24/7."),
], size=14)
rect(s, 6.9, 1.7, 5.85, 5.1, NAVY, MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, 7.2, 1.9, 5.3, 0.5, "Тип сайту", size=14, color=ORANGE, bold=True, font=HF)
text(s, 7.2, 2.3, 5.3, 0.9, "Корпоративний сайт-лендинг (односторінковий) сервісної компанії", size=18, color=WHITE, bold=True, font=HF)
bullets(s, 7.2, 3.35, 5.3, 3.4, [
    ("Мета:", "залучення B2B-заявок (лідогенерація) та самообслуговування клієнтів."),
    ("Інтерактив:", "калькулятор вартості, трекінг вантажу, форма заявки з валідацією."),
    ("Технології:", "HTML5 + CSS3 + JavaScript, згенеровано з допомогою ШІ, хостинг — GitHub Pages."),
], size=14, color=WHITE)

# ================= 3. Характеристика типу =================
s = prs.slides.add_slide(BLANK)
header(s, "Корпоративний сайт логістичної компанії", "Характеристика типу сайтів", 3)
text(s, 0.6, 1.55, 12.1, 0.9, "Корпоративний сайт — це «цифровий офіс» компанії: презентує послуги, формує довіру та конвертує відвідувача в клієнта. "
     "Для логістики він одночасно є робочим інструментом: клієнти відстежують вантажі й рахують вартість без дзвінка менеджеру.", size=15, color=MUTED)
cards = [("🎯", "Мета", "Генерація заявок, імідж, зменшення навантаження на кол-центр за рахунок онлайн-сервісів."),
         ("🧩", "Обов’язкові блоки", "Трекінг, калькулятор, послуги, тарифи, відділення/маршрути, контакти, FAQ."),
         ("🤝", "Тригери довіри", "Цифри компанії, відгуки, ліцензії/страхування, логотипи клієнтів, фото автопарку."),
         ("📱", "Технічні вимоги", "Адаптивність, швидкість, чіткий CTA у шапці, онлайн-чат/месенджери, SEO.")]
for i, (ic, h, b) in enumerate(cards):
    x = 0.6 + i * 3.08
    rect(s, x, 2.75, 2.9, 3.0, SOFT, MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, x + 0.25, 2.9, 1, 0.7, ic, size=30)
    text(s, x + 0.25, 3.6, 2.5, 0.5, h, size=15, color=NAVY, bold=True, font=HF)
    text(s, x + 0.25, 4.1, 2.45, 1.6, b, size=13, color=MUTED)
rect(s, 0.6, 6.0, 12.1, 0.8, NAVY, MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, 0.85, 6.03, 11.7, 0.74, [[("Різновиди: ", {"bold": True, "color": ORANGE}),
     ("сайти поштових операторів (B2C+B2B), сайти вантажних перевізників (B2B), біржі вантажів, лендинги експедиторів.", {"color": WHITE})]],
     size=14, anchor=MSO_ANCHOR.MIDDLE)

# ================= 4–8. Приклади =================
examples = [
    ("Нова пошта", "https://novaposhta.ua/", "c_novaposhta.png", "Лідер ринку · B2C + B2B",
     ["Hero-банер із відео та великим полем трекінгу — головна дія на першому екрані",
      "Мегаменю за сценаріями: Відправити / Отримати / Бізнесу / Міжнародна",
      "Кнопка «Створити відправлення», пошук відділень, вхід в особистий кабінет",
      "Фірмова червона гама + чат-бот «Потрібна допомога?»"],
     "Сильний бренд, все в 1 кліку; але сайт перевантажений B2C-сценаріями — бізнес-клієнту складно знайти вантажні тарифи."),
    ("Delivery (Делівері)", "https://www.delivery-auto.com/uk-UA", "c_delivery.png", "B2B · палетні та збірні вантажі",
     ["Яскрава жовта гама, слайдер з акціями на першому екрані",
      "Віджет-вкладки: трекінг · розрахунок вартості · виклик кур’єра · представництва",
      "Телефони гарячої лінії та «Зворотний дзвінок» у шапці",
      "Реєстрація / вхід для корпоративних клієнтів"],
     "Зручні вкладки-віджети; проте візуально застарілий дизайн, банер займає пів екрана, а калькулятор захований у вкладці."),
    ("SAT — логістика для бізнесу", "https://www.sat.ua/", "c_sat.png", "B2B · вантажі від 1 кг до 20 т",
     ["Нестандартна навігація: великі текстові пункти меню зліва",
      "Поле «номер вашого відправлення» — над слайдером",
      "Синьо-жовта гама, месенджери Telegram/Viber у шапці",
      "Банер про індивідуальні тарифи та персональний супровід"],
     "Чітке B2B-позиціонування; мінуси — каруселька з 10 слайдів, cookie-банер перекриває меню, мало інформації без скролу."),
    ("Meest", "https://ua.meest.com/", "c_meest.png", "Міжнародна доставка · B2C + B2B",
     ["Бічна панель з іконками швидких дій (відділення, тарифи, трекінг)",
      "Єдиний рядок «Пошук або трекінг» у шапці",
      "Перемикач «Приватним клієнтам / Для бізнесу» в блоці послуг",
      "Месенджери та чат-віджет, синьо-червона фірмова гама"],
     "Добре розділені аудиторії; але головний екран віддано промо Shopping, а не логістиці, — для B2B-клієнта це «шум»."),
    ("Raben Group Україна", "https://ukraine.raben-group.com/", "c_raben.png", "Європейський 3PL-оператор · B2B",
     ["Hero-фото вантажівки на трасі + червона плашка бренду",
      "Пошук відправлення та вхід до кабінету myRaben під банером",
      "Блок «Запитати ціну»: разове / регулярне перевезення",
      "Строге меню: Послуги, ІТ-рішення, Галузі, Контакти"],
     "Корпоративний, «європейський» стиль і сегментація за галузями; але немає миттєвого калькулятора — тільки запит ціни."),
]
for idx, (name, url, img, tag, elems, verdict) in enumerate(examples):
    n = 4 + idx
    s = prs.slides.add_slide(BLANK)
    header(s, name, f"Приклад {idx + 1} з 5 · аналіз сайту", n)
    _, pw, ph = picture(s, SC / img, 0.6, 1.6, w=7.3)
    url_bar(s, 0.6, 1.6, pw, url)
    tb = text(s, 0.6, 1.6 + ph + 0.12, 7.3, 0.35, [[("🔗 ", {}), (url, {"color": ORANGE, "bold": True})]], size=13)
    tb.text_frame.paragraphs[0].runs[1].hyperlink.address = url
    rect(s, 8.25, 1.6, 4.5, 0.45, NAVY, MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, 8.25, 1.6, 4.5, 0.45, tag, size=13, color=WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 8.25, 2.2, 4.5, 0.4, "Ключові елементи", size=16, color=NAVY, bold=True, font=HF)
    bullets(s, 8.25, 2.6, 4.5, 3.0, elems, size=13)
    rect(s, 8.25, 5.45, 4.5, 1.45, SOFT, MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, 8.25, 5.45, 0.08, 1.45, ORANGE)
    text(s, 8.45, 5.5, 4.25, 1.38, [[("Оцінка: ", {"bold": True, "color": NAVY}), (verdict, {})]], size=12, color=TEXT)

# ================= 9. Висновки =================
s = prs.slides.add_slide(BLANK)
header(s, "Висновки з аналізу сайтів-аналогів", "Узагальнення", 9)
blocks = [
    ("👥 Цільова аудиторія", ["Бізнес-відправники: виробники, дистриб’ютори, інтернет-магазини",
                             "Логісти та закупівельники компаній (25–50 р.)", "Імпортери/експортери, що працюють з ЄС",
                             "Приватні клієнти — у великих операторів (НП, Meest)"]),
    ("🎨 Кольорова гама", ["Фірмовий яскравий акцент: червоний (НП, Raben), жовтий (Delivery, SAT), синій (Meest, SAT)",
                          "Нейтральний білий/сірий фон для читабельності", "Контрастні CTA-кнопки «Відстежити», «Розрахувати»",
                          "Фото вантажівок, доріг та складів"]),
    ("🧩 Типові елементи інтерфейсу", ["Поле трекінгу на першому екрані — у 5 з 5 сайтів", "Калькулятор вартості / «Запитати ціну»",
                                       "Слайдер з акціями, мегаменю, пошук відділень", "Вхід в особистий кабінет, чат-бот, месенджери"]),
    ("🗂 Структура сторінок", ["Шапка: логотип · меню · пошук/трекінг · телефон · вхід", "Hero з головною дією → сервісні віджети → послуги",
                               "Цифри компанії, новини/акції, відгуки", "Підвал: контакти, соцмережі, юридична інформація, ©"]),
]
for i, (h, items) in enumerate(blocks):
    x = 0.6 + (i % 2) * 6.15; y = 1.6 + (i // 2) * 2.7
    rect(s, x, y, 5.95, 2.5, SOFT, MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, x + 0.25, y + 0.12, 5.5, 0.45, h, size=17, color=NAVY, bold=True, font=HF)
    bullets(s, x + 0.25, y + 0.6, 5.5, 1.85, items, size=12.5)

# ================= 10. Чим кращий (1) =================
s = prs.slides.add_slide(BLANK)
header(s, "Чим мій сайт буде кращим (1/2): унікальність", "Переваги OZK Logistics", 10)
items = [
    ("Прозорий калькулятор одразу на сторінці:", "FTL за км (22–28 ₴) або LTL за палету (600–800 ₴), опції «терміново», «митниця», «знижка постійного клієнта» — ціна миттєво, без реєстрації."),
    ("Трекінг з таймлайном етапів:", "прогрес-бар, статус на кордоні, прогноз доставки — а не лише текстовий статус."),
    ("Фокус тільки на B2B:", "без B2C-«шуму», промо-слайдерів і зайвих розділів — усе для логіста компанії."),
    ("Комунікація про енергонезалежність:", "генератори + Starlink — диспетчерська 24/7 під час блекаутів (жоден з аналогів цього не підкреслює)."),
    ("Персональний логіст 24/7 і фіксований тариф", "у договорі — головні болі клієнтів великих операторів."),
]
bullets(s, 0.6, 1.65, 6.2, 5.2, items, size=14, marker="✔")
_, pw, ph = picture(s, SC / "own_calc.png", 7.1, 1.6, w=5.65)
url_bar(s, 7.1, 1.6, pw, SITE_URL)
text(s, 7.1, 1.6 + ph + 0.1, 5.65, 0.4, "Онлайн-калькулятор вартості на створеному сайті", size=11, color=MUTED, align=PP_ALIGN.CENTER)
rect(s, 7.1, 6.0, 5.65, 0.85, NAVY, MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, 7.3, 6.0, 5.3, 0.85, [[("Слоган: ", {"bold": True, "color": ORANGE}), ("«Доставимо ваш вантаж вчасно — за тарифом, зафіксованим заздалегідь»", {"color": WHITE})]],
     size=13, anchor=MSO_ANCHOR.MIDDLE)

# ================= 11. Чим кращий (2) =================
s = prs.slides.add_slide(BLANK)
header(s, "Чим мій сайт буде кращим (2/2): порівняння", "Переваги OZK Logistics", 11)
rows = [("Критерій", "OZK", "Нова пошта", "Delivery", "SAT", "Meest", "Raben"),
        ("Трекінг на 1-му екрані", "✔", "✔", "✔", "✔", "✔", "✔"),
        ("Калькулятор без переходу на іншу сторінку", "✔", "—", "±", "—", "—", "—"),
        ("Таймлайн етапів доставки з прогресом", "✔", "±", "—", "—", "—", "—"),
        ("Фокус лише на B2B, без промо-шуму", "✔", "—", "±", "✔", "—", "✔"),
        ("Інформація про роботу в блекаут", "✔", "—", "—", "—", "—", "—"),
        ("Порівняльна таблиця тарифів з ринком", "✔", "—", "—", "—", "—", "—"),
        ("Адаптивність + легкість (без фреймворків)", "✔", "±", "±", "±", "±", "±")]
tbl = s.shapes.add_table(len(rows), len(rows[0]), Inches(0.6), Inches(1.6), Inches(7.6), Inches(3.9)).table
widths = [3.0] + [0.77] * 6
for j, w in enumerate(widths): tbl.columns[j].width = Inches(w)
for i, r in enumerate(rows):
    for j, v in enumerate(r):
        c = tbl.cell(i, j); c.text = v
        p = c.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
        f = p.runs[0].font; f.size = Pt(11 if i else 11); f.name = BF; f.bold = i == 0 or j == 1
        c.vertical_anchor = MSO_ANCHOR.MIDDLE
        c.fill.solid()
        if i == 0: c.fill.fore_color.rgb = ORANGE if j == 1 else NAVY; f.color.rgb = WHITE
        elif j == 1: c.fill.fore_color.rgb = RGBColor(0xFF, 0xF1, 0xE5); f.color.rgb = GREEN
        else:
            c.fill.fore_color.rgb = WHITE if i % 2 else SOFT
            f.color.rgb = GREEN if v == "✔" else (RED if v == "—" else MUTED)
text(s, 0.6, 5.6, 7.6, 0.4, "✔ — є; ± — частково; — — відсутнє (за результатами аналізу головних сторінок, листопад 2026)", size=10, color=MUTED)
bullets(s, 0.6, 6.0, 7.6, 1.0, [("Сучасний дизайн:", "темно-синя + помаранчева гама (надійність + енергія), анімації, мобільна версія.")], size=13)
_, pw, ph = picture(s, SC / "own_mobile.png", 8.6, 1.6, h=5.2, border=False)
text(s, 8.6 + pw + 0.2, 1.6, 12.75 - (8.6 + pw + 0.2), 5.2,
     [[("Мобільна версія", {"bold": True, "color": NAVY, "size": 13})],
      "бургер-меню, адаптивна сітка, плаваюча кнопка швидкої заявки.",
      [("12 блоків сайту:", {"bold": True, "color": NAVY, "size": 13})],
      "шапка, hero, цифри, послуги, калькулятор, трекінг, процес, переваги + таблиця порівняння, маршрути, відгуки, FAQ, заявка, підвал з ©."],
     size=11, color=MUTED)

# ================= 12. Посилання =================
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.333, 7.5, NAVY)
text(s, 0.8, 0.5, 10, 0.4, "РЕЗУЛЬТАТ", size=13, color=ORANGE, bold=True, font=HF)
text(s, 0.8, 0.85, 11, 0.9, "Інтернет-покликання на створений сайт", size=32, color=WHITE, bold=True, font=HF)
rect(s, 0.8, 1.95, 7.2, 1.0, WHITE, MSO_SHAPE.ROUNDED_RECTANGLE)
tb = text(s, 1.0, 1.95, 6.9, 1.0, [[("🌐 ", {}), (SITE_URL, {"bold": True, "color": NAVY})]], size=20, anchor=MSO_ANCHOR.MIDDLE)
tb.text_frame.paragraphs[0].runs[1].hyperlink.address = SITE_URL
tb = text(s, 0.8, 3.1, 7.2, 0.4, [[("Репозиторій: ", {"color": LIGHT}), (REPO_URL, {"color": ORANGE})]], size=14)
tb.text_frame.paragraphs[0].runs[1].hyperlink.address = REPO_URL
bullets(s, 0.8, 3.65, 7.2, 3.0, [
    ("Спосіб створення:", "штучний інтелект (генерація HTML/CSS/JS за промптами)"),
    ("Хостинг:", "GitHub Pages (Branch: main, Folder: /root)"),
    ("Файли:", "index.html, style.css, script.js"),
    ("Підвал сайту:", "© Олександр Захарчук 2026, o.zakharchuk@knute.edu.ua"),
], size=14, color=WHITE)
qr = qrcode.QRCode(border=2, box_size=10); qr.add_data(SITE_URL); qr.make(fit=True)
qr.make_image(fill_color="#0B1F3A", back_color="white").save(SC / "qr.png")
s.shapes.add_picture(str(SC / "qr.png"), Inches(8.6), Inches(1.95), Inches(1.9), Inches(1.9))
text(s, 10.65, 2.2, 2.1, 1.4, "Скануйте QR-код, щоб відкрити сайт на смартфоні", size=12, color=LIGHT)
_, pw, ph = picture(s, SC / "own_hero.png", 8.6, 4.1, w=4.15, border=False)
text(s, 0.8, 6.4, 6, 0.55, "Дякую за увагу!", size=20, color=ORANGE, bold=True, font=HF)
footer(s, 12, dark=True)

out = ROOT / "Презентація_ЛР5_Захарчук_OZK_Logistics.pptx"
prs.save(out)
print("saved", out, "url:", SITE_URL)
