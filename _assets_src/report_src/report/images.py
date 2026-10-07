# -*- coding: utf-8 -*-
"""Подготовка картинок для отчёта: обрезка длинных скриншотов, сборка сводных картинок, схема сайта.
Исходники в report_assets не меняются, результат лежит в report/build/img."""
import os
import textwrap

from PIL import Image, ImageDraw, ImageFont

ASSETS = '/home/claude/v2/report_assets'
SCREENS = os.path.join(ASSETS, 'screens')
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'build', 'img')
GREY = (70, 70, 76)


def fpath(bold=False, mono=False):
    if mono:
        return '/usr/share/fonts/truetype/liberation/LiberationMono-%s.ttf' % ('Bold' if bold else 'Regular')
    return '/usr/share/fonts/truetype/liberation/LiberationSans-%s.ttf' % ('Bold' if bold else 'Regular')


def F(size, bold=False, mono=False):
    return ImageFont.truetype(fpath(bold, mono), size)


def load(folder, name):
    return Image.open(os.path.join(ASSETS, folder, name)).convert('RGB')


def save(im, name):
    im.save(os.path.join(OUT, name), optimize=True)


def row(items, name, gap=24, bg=GREY, height=None):
    """Картинки в ряд (выравнивание по верху)."""
    ims = []
    for im in items:
        if height:
            im = im.resize((round(im.width * height / im.height), height), Image.LANCZOS)
        ims.append(im)
    W = sum(i.width for i in ims) + gap * (len(ims) - 1)
    H = max(i.height for i in ims)
    sheet = Image.new('RGB', (W, H), bg)
    x = 0
    for i in ims:
        sheet.paste(i, (x, 0))
        x += i.width + gap
    save(sheet, name)


def col(items, name, gap=16, bg=GREY):
    """Картинки друг под другом, по левому краю."""
    W = max(i.width for i in items)
    H = sum(i.height for i in items) + gap * (len(items) - 1)
    sheet = Image.new('RGB', (W, H), bg)
    y = 0
    for i in items:
        sheet.paste(i, (0, y))
        y += i.height + gap
    save(sheet, name)


# ---------------------------------------------------------------- схема структуры сайта
def sitemap():
    W, H = 2000, 1180
    im = Image.new('RGB', (W, H), 'white')
    d = ImageDraw.Draw(im)
    RED, BLK = (215, 20, 10), (17, 17, 17)

    # шапка и подвал
    d.rounded_rectangle((40, 30, W - 40, 150), 14, fill=(17, 17, 17), outline=RED, width=5)
    d.text((70, 44), 'Общая шапка (на всех страницах)', font=F(32, True), fill='white')
    d.text((70, 94), 'логотип GM GYM  ·  Главная  ·  Направления  ·  БЖЖ  ·  Тренеры  ·  Контакты  ·  кнопка «Записаться»',
           font=F(30), fill=(230, 230, 230))

    def box(x, y, w, h, title, sub, lines, fill, stroke):
        d.rounded_rectangle((x, y, x + w, y + h), 18, fill=fill, outline=stroke, width=5)
        d.text((x + 24, y + 18), title, font=F(40, True), fill=BLK)
        d.text((x + 24, y + 70), sub, font=F(27), fill=(80, 80, 80))
        yy = y + 125
        for ln in lines:
            d.text((x + 24, yy), '–  ' + ln, font=F(29), fill=(30, 30, 30))
            yy += 42

    box(40, 230, 600, 500, 'index.html', 'Главная', [
        'первый экран', 'о клубе и удобства', 'направления (8 карточек)', 'тренеры (8 карточек)',
        'наш зал (3 фото)', 'видео', 'призыв «Записаться»'], (236, 243, 255), (40, 90, 170))
    box(740, 230, 600, 500, 'bjj.html', 'Страница объекта: БЖЖ', [
        'первый экран', 'что такое БЖЖ, кому подойдёт', 'расписание (таблица)', 'тренеры по БЖЖ (5)',
        'что взять на тренировку', 'форма записи  #signup'], (236, 243, 255), (40, 90, 170))
    box(1440, 230, 520, 500, 'thanks.html', 'Заявка отправлена', [
        'заголовок и текст', 'кнопка «На главную»', 'кнопка «К расписанию БЖЖ»'], (236, 250, 238), (40, 130, 60))

    def arrow(p1, p2, color=(200, 40, 20), w=6, head=True):
        import math
        d.line([p1, p2], fill=color, width=w)
        if not head:
            return
        ang = math.atan2(p2[1] - p1[1], p2[0] - p1[0])
        L = 30
        d.polygon([p2, (p2[0] - L * math.cos(ang - 0.45), p2[1] - L * math.sin(ang - 0.45)),
                   (p2[0] - L * math.cos(ang + 0.45), p2[1] - L * math.sin(ang + 0.45))], fill=color)

    def badge(x, y, n):
        d.ellipse((x - 22, y - 22, x + 22, y + 22), fill='white', outline=(200, 40, 20), width=5)
        d.text((x, y), str(n), font=F(30, True), fill=(200, 40, 20), anchor='mm')

    # 1 главная -> БЖЖ
    arrow((640, 330), (740, 330)); badge(690, 295, 1)
    # 2 БЖЖ -> спасибо (submit)
    arrow((1340, 330), (1440, 330)); badge(1390, 295, 2)
    # 5 БЖЖ -> главная (меню)
    arrow((740, 560), (640, 560)); badge(690, 595, 5)
    # 3 спасибо -> главная (внизу), 4 спасибо -> БЖЖ
    arrow((1700, 730), (1700, 850), head=False); arrow((1700, 850), (340, 850), head=False)
    arrow((340, 850), (340, 730))
    badge(1700, 800, 3)
    arrow((1560, 730), (1560, 790), head=False); arrow((1560, 790), (1040, 790), head=False)
    arrow((1040, 790), (1040, 730))
    badge(1560, 760, 4)

    # подвал
    d.rounded_rectangle((40, 900, W - 40, 970), 14, fill=(240, 240, 240), outline=(90, 90, 90), width=4)
    d.text((70, 920), 'Общий подвал: адрес и метро, телефон, WhatsApp, e-mail, ВКонтакте, меню, пометка «учебный проект»',
           font=F(28), fill=(40, 40, 40))
    # легенда
    legend = [
        '1  меню «БЖЖ», карточка БЖЖ, кнопки «Записаться» (ведут на bjj.html, #signup)',
        '2  кнопка «Отправить заявку»: form action="thanks.html" method="get" (только если поля заполнены верно)',
        '3  кнопка «На главную»',
        '4  кнопка «К расписанию БЖЖ»',
        '5  меню «Главная», «Направления», «Тренеры», «Контакты»',
    ]
    y = 995
    for ln in legend:
        d.text((60, y), ln, font=F(27), fill=(30, 30, 30))
        y += 36
    save(im, 'sitemap.png')


# ---------------------------------------------------------------- лог git как картинка
def gitlog():
    """Вывод git log --graph в две колонки (чтобы картинка была невысокой)."""
    with open(os.path.join(ASSETS, 'git_log.txt'), encoding='utf-8') as f:
        raw = f.read().rstrip('\n').split('\n')
    lines = []
    for ln in raw:
        if len(ln) > 62:
            lines.extend(textwrap.wrap(ln, 62, subsequent_indent='|' + ' ' * 11, break_long_words=False))
        else:
            lines.append(ln)
    # граница колонок: ближайшая к середине строка-слияние вида «*   hash Слияние ...»
    mid = len(lines) // 2
    cut = min((i for i, l in enumerate(lines) if l.startswith('*   ')), key=lambda i: abs(i - mid))
    cols = [lines[:cut], lines[cut:]]
    fs, lh = 20, 26
    font = F(fs, mono=True)
    cw = font.getlength('M')
    colw = int(cw * 64) + 20
    H = lh * max(len(c) for c in cols) + 36
    W = colw * 2 + 20
    im = Image.new('RGB', (W, H), (250, 250, 250))
    d = ImageDraw.Draw(im)
    for k, c in enumerate(cols):
        y = 18
        for ln in c:
            d.text((14 + k * (colw + 10), y), ln, font=font, fill=(25, 25, 25))
            y += lh
    d.line((colw + 10, 10, colw + 10, H - 10), fill=(170, 170, 170), width=2)
    d.rectangle((0, 0, W - 1, H - 1), outline=(150, 150, 150), width=2)
    save(im, 'git_log_render.png')


def tile_grid(files, labels, name, tw=800, gap=16):
    """Сетка 2 x N из скриншотов аналогов; на каждом снизу плашка с названием сайта."""
    tiles = []
    for f, lab in zip(files, labels):
        im = load('analogs', f)
        im = im.resize((tw, round(im.height * tw / im.width)), Image.LANCZOS)
        d = ImageDraw.Draw(im)
        d.rectangle((0, im.height - 44, tw, im.height), fill=(17, 17, 17))
        d.text((16, im.height - 22), lab, font=F(26, True), fill='white', anchor='lm')
        tiles.append(im)
    th = tiles[0].height
    rows = (len(tiles) + 1) // 2
    sheet = Image.new('RGB', (tw * 2 + gap, th * rows + gap * (rows - 1)), (255, 255, 255))
    for i, t in enumerate(tiles):
        sheet.paste(t, ((i % 2) * (tw + gap), (i // 2) * (th + gap)))
    save(sheet, name)


def labeled(im, text, size=30):
    """Подпись над картинкой (для сводных рисунков)."""
    bar = 46
    out = Image.new('RGB', (im.width, im.height + bar), (255, 255, 255))
    d = ImageDraw.Draw(out)
    d.text((8, bar // 2), text, font=F(size, True), fill=(0, 0, 0), anchor='lm')
    out.paste(im, (0, bar))
    return out


def prepare_images():
    os.makedirs(OUT, exist_ok=True)
    S = lambda n: load('screens', n)

    # аналоги: сетки 2 x 2
    tile_grid(['gmgym.png', 'bjj77.png', 'ares_bjj.png', 'mmadojo.png'],
              ['а) gmgym.ru', 'б) bjj77.ru (Титан BJJ)', 'в) aresfight.ru (Арес)', 'г) mmadojo.ru (MMA Dojo)'], 'an_dom.png')
    tile_grid(['graciebarra.png', 'evolve.png', 'att.png'],
              ['а) graciebarra.com', 'б) evolve-mma.com', 'в) americantopteam.com'], 'an_for.png')

    # шапка и подвал (ПК)
    col([labeled(S('header_1440.png').crop((100, 0, 1340, 67)), 'а) шапка', 26),
         labeled(S('footer_1440.png').crop((100, 0, 1340, 404)), 'б) подвал', 26)], 'hdrftr.png', gap=8, bg=(255, 255, 255))
    vb = S('video_block.png').crop((100, 60, 1340, 780))
    hl = S('index_1440.png').crop((100, 2974, 1340, 3505))
    row([labeled(hl, 'а) раздел «Наш зал»', 30), labeled(vb, 'б) раздел «Видео о клубе»', 30)], 'media_pair.png', gap=24,
        bg=(255, 255, 255))

    # главная и БЖЖ рядом: первый экран и следующие блоки
    row([labeled(S('index_1440.png').crop((0, 0, 1440, 1900)), 'а) index.html', 44),
         labeled(S('bjj_1440.png').crop((0, 0, 1440, 1900)), 'б) bjj.html', 44)], 'pages_top.png', gap=30, bg=(255, 255, 255))

    # три раскладки
    a = S('index_1440.png')
    pc = Image.new('RGB', (1440, 518 + 16 + 400), GREY)
    pc.paste(a.crop((0, 0, 1440, 518)), (0, 0))
    pc.paste(a.crop((0, 1279, 1440, 1679)), (0, 534))
    save(pc, 'lay_pc.png')
    b = S('index_768.png')
    row([b.crop((0, 0, 768, 824)), b.crop((0, 1637, 768, 2400))], 'lay_tab.png', gap=20)
    c = S('index_375.png')
    m = S('schedule_mobile.png')
    mp = Image.new('RGB', (375, 900), (244, 244, 244))
    mp.paste(m, (0, 0))
    row([c.crop((0, 0, 375, 900)), c.crop((0, 2073, 375, 2973)), c.crop((0, 3524, 375, 4424)), mp], 'lay_mob.png', gap=16)

    # форма: пустая и заполненная рядом
    row([labeled(S('form_empty_submit.png').crop((130, 330, 930, 670)), 'а) пустая форма', 26),
         labeled(S('form_filled.png').crop((130, 330, 930, 670)), 'б) заполненная форма', 26)], 'form_pair.png', gap=20,
        bg=(255, 255, 255))
    save(S('thanks_after_submit.png').crop((0, 0, 1440, 700)), 'thanks_crop.png')
    save(S('cards_hover.png').crop((100, 300, 1340, 740)), 'hover_crop.png')

    # онлайн-проверки: три Nu в одной картинке, оставляем заголовок, начало кода и зелёную полосу
    parts = []
    for n, lab in (('index', 'а) index.html'), ('bjj', 'б) bjj.html'), ('thanks', 'в) thanks.html')):
        im = load('online', 'nu_%s.png' % n)
        blk = Image.new('RGB', (1280, 45 + 105 + 45), 'white')
        blk.paste(im.crop((0, 95, 1280, 140)), (0, 0))
        blk.paste(im.crop((0, 240, 1280, 345)), (0, 45))
        blk.paste(im.crop((0, 565, 1280, 610)), (0, 150))
        parts.append(labeled(blk, lab, 30))
    col(parts, 'nu_all.png', gap=10, bg=(255, 255, 255))
    css = load('online', 'css_style.png')
    top = css.crop((0, 190, 1280, 300))
    warn = css.crop((0, 1170, 1280, 1720))
    col([top, warn], 'css_val.png', gap=10, bg=(255, 255, 255))

    col([S('skip_link.png'), S('focus_visible.png')], 'a11y_pair.png', gap=12)
    sitemap()
    gitlog()
    return OUT


if __name__ == '__main__':
    print(prepare_images())
