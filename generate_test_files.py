#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор эталонных тестовых презентаций для проекта «Проверка .pptx перед уроком».

Создаёт в папке test_files/:
  - broken.pptx           — презентация с полным набором типовых дефектов
  - clean.pptx            — «здоровая» презентация со сложной реальной вёрсткой
                            (фоновая подложка под текстом, зеркальная картинка flipH,
                            прозрачная PNG-рамка поверх текста, шрифт темы +mn-lt,
                            микросмещение -24 EMU, ссылка mailto:), на которой
                            НЕ должно быть ни одного ложного срабатывания
  - expected_defects.json — эталон для самопроверки (--test)

Запуск:  python3 generate_test_files.py
"""

import json
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

# Windows: не падать на кириллице в консоли (cp1252/cp866)
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

BASE = Path(__file__).resolve().parent
OUT = BASE / "test_files"
ASSETS = OUT / "assets"


# ---------------------------------------------------------------- вспомогательные

def _get_font(size: int):
    """Найти шрифт для генерации тестовых изображений."""
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/TTF/DejaVuSans.ttf",
        "C:/Windows/Fonts/arial.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
    ]
    for p in candidates:
        if Path(p).exists():
            return ImageFont.truetype(p, size)
    return ImageFont.load_default()


def make_scheme(path: Path) -> None:
    """Нарисовать чёткую учебную схему 900x600 px."""
    w, h = 900, 600
    img = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(img)
    box = dict(outline="navy", width=5)
    d.rectangle([50, 70, 380, 250], **box)
    d.rectangle([520, 70, 850, 250], **box)
    d.rectangle([285, 380, 615, 560], outline="darkgreen", width=5)
    d.line([380, 160, 510, 160], fill="black", width=5)
    d.polygon([(520, 160), (495, 147), (495, 173)], fill="black")
    d.line([230, 250, 340, 375], fill="black", width=5)
    d.polygon([(350, 385), (323, 370), (339, 351)], fill="black")
    d.line([670, 250, 560, 375], fill="black", width=5)
    d.polygon([(550, 385), (577, 351), (561, 370)], fill="black")
    f = _get_font(40)
    d.text((105, 150), "Процесс А", fill="black", font=f)
    d.text((575, 150), "Процесс Б", fill="black", font=f)
    d.text((345, 465), "Результат", fill="black", font=f)
    img.save(path)


def make_tiny(path: Path) -> None:
    """Крошечная картинка 80x80 px — имитация растянутой иконки из интернета."""
    img = Image.new("RGB", (80, 80), (255, 220, 120))
    d = ImageDraw.Draw(img)
    d.ellipse([10, 10, 70, 70], fill="orange", outline="red", width=3)
    d.text((24, 30), "?!?", fill="black", font=_get_font(22))
    img.save(path)


def make_soft_bg(path: Path) -> None:
    """Светлая фоновая подложка 1600x900 px для проверки отсутствия ложных срабатываний."""
    img = Image.new("RGB", (1600, 900), (242, 247, 255))
    d = ImageDraw.Draw(img)
    d.rectangle([20, 20, 1580, 880], outline=(190, 210, 240), width=8)
    img.save(path)


def make_transparent_frame(path: Path) -> None:
    """PNG-рамка с полностью прозрачной серединой (проверка учёта альфа-канала)."""
    img = Image.new("RGBA", (800, 500), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, 799, 499], outline=(40, 110, 180, 255), width=14)
    img.save(path)


def add_text(slide, left, top, width, height, text, size=24,
             font_name="Arial", bold=False):
    """Добавить текстовое поле и задать шрифт всем абзацам."""
    box = slide.shapes.add_textbox(Inches(left), Inches(top),
                                   Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    tf.text = text
    for para in tf.paragraphs:
        for run in para.runs:
            run.font.size = Pt(size)
            if font_name is not None:
                run.font.name = font_name
            run.font.bold = bold
    return box


def set_flip(pic, attr: str = "flipV") -> None:
    """Установить flipV=1 или flipH=1 в XML картинки."""
    xfrm = pic._element.spPr.find(qn("a:xfrm"))
    if xfrm is None:
        raise RuntimeError("у картинки нет xfrm")
    xfrm.set(attr, "1")


def set_auto_advance(slide, ms: int) -> None:
    """Включить автопереключение слайда по таймеру."""
    sld = slide._element
    transition = sld.makeelement(
        qn("p:transition"), {"advClick": "0", "advTm": str(ms)}
    )
    transition.append(transition.makeelement(qn("p:fade"), {}))
    clr = sld.find(qn("p:clrMapOvr"))
    anchor = clr if clr is not None else sld.find(qn("p:cSld"))
    anchor.addnext(transition)


def new_presentation() -> Presentation:
    prs = Presentation()
    prs.slide_width = Inches(13.333)   # 16:9 (33,9 x 19,1 см)
    prs.slide_height = Inches(7.5)
    return prs


# ---------------------------------------------------------------- сборка файлов

def build_broken(path: Path, scheme: Path, tiny: Path) -> list:
    prs = new_presentation()
    blank = prs.slide_layouts[6]
    expected = []

    # --- слайд 1: обложка (дефектов нет) ---
    s = prs.slides.add_slide(blank)
    add_text(s, 1.5, 2.5, 10.5, 1.6,
             "Тестовая презентация «СЛОМАННАЯ»\n"
             "Набор эталонных дефектов для самопроверки", 32, bold=True)
    add_text(s, 1.5, 4.5, 10.5, 1.0,
             "Эталон для команды python check_presentation.py --test", 20)

    # --- слайд 2: перевёрнутая вверх ногами картинка ---
    s = prs.slides.add_slide(blank)
    add_text(s, 0.5, 0.3, 12, 0.9, "Слайд 2. Схема процесса", 28, bold=True)
    pic = s.shapes.add_picture(str(scheme), Inches(3.5), Inches(1.6),
                               Inches(6), Inches(4))
    set_flip(pic, "flipV")
    expected.append({
        "slide": 2,
        "defect": "flipped_image",
        "description": "картинка перевёрнута вверх ногами (flipV=1)",
    })

    # --- слайд 3: объекты вылезают за границы слайда ---
    s = prs.slides.add_slide(blank)
    add_text(s, 0.5, 0.3, 12, 0.9, "Слайд 3. Выход за края слайда", 28, bold=True)
    add_text(s, 0.7, 5.6, 8, 3.2,
             "Этот текстовый блок начинается слишком низко и заканчивается "
             "значительно ниже нижней границы слайда. В режиме показа последние "
             "строки просто обрезаются проектором.", 22)
    expected.append({
        "slide": 3,
        "defect": "object_out_of_slide",
        "description": "текстовое поле выходит за нижнюю границу слайда (8.8\" > 7.5\")",
    })
    s.shapes.add_picture(str(scheme), Inches(10.5), Inches(1.6),
                         Inches(4), Inches(2.7))
    expected.append({
        "slide": 3,
        "defect": "object_out_of_slide",
        "description": "картинка выходит за правый край слайда (14.5\" > 13.333\")",
    })

    # --- слайд 4: непрозрачная картинка перекрывает текст сверху ---
    s = prs.slides.add_slide(blank)
    add_text(s, 0.5, 0.3, 12, 0.9, "Слайд 4. Важные определения", 28, bold=True)
    add_text(s, 0.8, 1.6, 6, 3.4,
             "Здесь находится текст определения. Поверх него случайно положили "
             "непрозрачную схему на верхний слой — правая половина строк закрыта "
             "изображением.", 20)
    s.shapes.add_picture(str(scheme), Inches(4.5), Inches(2.6),
                         Inches(6), Inches(4))
    expected.append({
        "slide": 4,
        "defect": "overlap_picture_text",
        "description": "непрозрачная картинка на верхнем слое закрывает текст",
    })

    # --- слайд 5: картинка низкого разрешения ---
    s = prs.slides.add_slide(blank)
    add_text(s, 0.5, 0.3, 12, 0.9, "Слайд 5. Растянутая картинка", 28, bold=True)
    add_text(s, 0.7, 1.4, 6, 2.5,
             "Картинка 80x80 пикселей растянута на пол-экрана (~13 DPI). "
             "На школьном проекторе она превратится в размытые квадраты.", 20)
    s.shapes.add_picture(str(tiny), Inches(7), Inches(1.6),
                         Inches(6), Inches(5.5))
    expected.append({
        "slide": 5,
        "defect": "low_resolution_image",
        "description": "80x80 px на 6\"x5.5\" -> ~13 DPI (< 75 DPI)",
    })

    # --- слайд 6: нестандартный веб-шрифт (не встроен в файл) ---
    s = prs.slides.add_slide(blank)
    add_text(s, 0.5, 0.3, 12, 0.9, "Слайд 6. Нестандартный шрифт", 28, bold=True)
    add_text(s, 0.7, 1.6, 11, 3.0,
             "Этот текст набран дизайнерским шрифтом Montserrat (часто остаётся "
             "после экспорта из Canva или онлайн-шаблонов), которого нет в "
             "стандартной поставке Windows и Office.", 22, font_name="Montserrat")
    expected.append({
        "slide": 6,
        "defect": "nonstandard_font",
        "description": "шрифт 'Montserrat' не входит в стандартный набор и не внедрён",
    })

    # --- слайд 7: забытый автопереход ---
    s = prs.slides.add_slide(blank)
    add_text(s, 0.5, 0.3, 12, 0.9, "Слайд 7. Автопереход по таймеру", 28, bold=True)
    add_text(s, 0.7, 1.6, 11, 2.5,
             "У этого слайда включено автопереключение через 2 секунды — "
             "слайд сам уедет вперёд посреди объяснения учителя.", 22)
    set_auto_advance(s, 2000)
    expected.append({
        "slide": 7,
        "defect": "auto_advance_transition",
        "description": "<p:transition advTm=\"2000\"> — сам переключится через 2 с",
    })

    # --- слайд 8: внешняя ссылка ---
    s = prs.slides.add_slide(blank)
    add_text(s, 0.5, 0.3, 12, 0.9, "Слайд 8. Внешняя интернет-ссылка", 28, bold=True)
    add_text(s, 0.7, 1.6, 11, 2.0,
             "Слайд ссылается на внешний ролик в интернете — в кабинете без сети "
             "ссылка не откроется.", 22)
    box = add_text(s, 0.7, 4.0, 6, 1.0, "Открыть видеоурок", 24)
    run = box.text_frame.paragraphs[0].runs[0]
    run.hyperlink.address = "https://example.com/video-lesson"
    expected.append({
        "slide": 8,
        "defect": "external_link",
        "description": "внешняя гиперссылка (TargetMode=External)",
    })

    # --- слайд 9: слишком мелкий нечитаемый шрифт ---
    s = prs.slides.add_slide(blank)
    add_text(s, 0.5, 0.3, 12, 0.9, "Слайд 9. Мелкий нечитаемый кегль", 28, bold=True)
    add_text(s, 0.7, 1.6, 11, 3.0,
             "Этот абзац скопирован мелким 10-м кеглем — с третьей парты "
             "ученики не смогут прочитать ни одного слова на проекторе.", 10)
    expected.append({
        "slide": 9,
        "defect": "small_font",
        "description": "кегль 10 pt (< 14 pt) нечитаем с задних парт",
    })

    prs.save(path)
    return expected


def build_clean(path: Path, scheme: Path, soft_bg: Path, frame_png: Path) -> None:
    """
    Здоровая презентация со сложными реальными приёмами вёрстки, которые
    НЕ должны вызывать ложных срабатываний анализатора.
    """
    prs = new_presentation()
    blank = prs.slide_layouts[6]

    # Слайд 1: обложка
    s = prs.slides.add_slide(blank)
    add_text(s, 1.5, 2.8, 10.5, 1.5,
             "Тестовая презентация «ЧИСТАЯ»\nЗдоровый файл без дефектов", 32, bold=True)

    # Слайд 2: фоновая подложка ПОД текстом (z=0), зеркальная схема (flipH=1)
    # и декоративная PNG-рамка с прозрачным центром ПОВЕРХ текста!
    s = prs.slides.add_slide(blank)
    s.shapes.add_picture(str(soft_bg), Inches(0), Inches(0),
                         Inches(13.333), Inches(7.5))
    add_text(s, 0.8, 0.4, 11.5, 0.9, "Фон снизу и прозрачная рамка сверху — это норма", 28, bold=True)
    add_text(s, 1.0, 2.0, 5.2, 2.8,
             "Текст расположен поверх фоновой подложки и внутри прозрачной "
             "PNG-рамки. Справа — отражённая по горизонтали (flipH) иллюстрация.", 20)
    s.shapes.add_picture(str(frame_png), Inches(0.8), Inches(1.7),
                         Inches(5.6), Inches(3.4))
    mirrored = s.shapes.add_picture(str(scheme), Inches(7.0), Inches(1.7),
                                    Inches(5.5), Inches(3.7))
    set_flip(mirrored, "flipH")

    # Слайд 3: служебный шрифт темы (+mn-lt), микропогрешность -24 EMU и почта mailto:
    s = prs.slides.add_slide(blank)
    title_box = add_text(s, 0.5, 0.0, 12, 0.9, "Шрифт темы и контакты автора", 28, bold=True)
    title_box.top = -24  # погрешность округления PowerPoint (-0,00067 мм)

    body = add_text(s, 0.8, 1.6, 11, 3.0,
                    "Этот текст использует стандартный токен шрифта темы +mn-lt, "
                    "который в PowerPoint ссылается на Calibri/Arial.", 22, font_name="+mn-lt")
    _ = body

    mail_box = add_text(s, 0.8, 5.2, 6, 0.9, "Написать учителю", 20)
    mail_box.text_frame.paragraphs[0].runs[0].hyperlink.address = "mailto:teacher@school.edu"

    prs.save(path)


def main() -> None:
    OUT.mkdir(exist_ok=True)
    ASSETS.mkdir(exist_ok=True)

    scheme = ASSETS / "scheme.png"
    tiny = ASSETS / "tiny.png"
    soft_bg = ASSETS / "soft_bg.png"
    frame_png = ASSETS / "frame_transparent.png"
    make_scheme(scheme)
    make_tiny(tiny)
    make_soft_bg(soft_bg)
    make_transparent_frame(frame_png)

    broken = OUT / "broken.pptx"
    clean = OUT / "clean.pptx"
    defects = build_broken(broken, scheme, tiny)
    build_clean(clean, scheme, soft_bg, frame_png)

    expected = {
        "broken.pptx": defects,
        "clean.pptx": [],
    }
    (OUT / "expected_defects.json").write_text(
        json.dumps(expected, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    print("Готово! Эталонные файлы обновлены:")
    for name in ("broken.pptx", "clean.pptx", "expected_defects.json"):
        f = OUT / name
        print(f"  {f.relative_to(BASE)}  ({f.stat().st_size} байт)")
    print(f"\nДефектов в эталоне broken.pptx: {len(defects)}")


if __name__ == "__main__":
    main()
