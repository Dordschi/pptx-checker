#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор тестовых презентаций для проекта «Проверка .pptx перед уроком».

Создаёт в папке test_files/:
  - broken.pptx        — презентация, где КАЖДЫЙ слайд содержит ровно один дефект
  - clean.pptx         — «здоровая» презентация (на ней проверяем, что нет ложных срабатываний)
  - expected_defects.json — эталон: в каком файле и на каком слайде какой дефект спрятан

Запуск:  python3 generate_test_files.py
Зависимости: pip install python-pptx pillow
"""

import json
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn

BASE = Path(__file__).resolve().parent
OUT = BASE / "test_files"
ASSETS = OUT / "assets"

# ---------------------------------------------------------------- вспомогательные

def _get_font(size: int):
    """Найти нормальный шрифт для рисования картинок (иначе встроенный мелкий)."""
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
    """Нарисовать простую «схему» с подписями — как в учебных презентациях."""
    w, h = 900, 600
    img = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(img)
    box = dict(outline="navy", width=5)
    d.rectangle([50, 70, 380, 250], **box)
    d.rectangle([520, 70, 850, 250], **box)
    d.rectangle([285, 380, 615, 560], outline="darkgreen", width=5)
    # стрелки
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
    """Крошечная картинка 80x80 — имитация «пиксельной» картинки с интернета."""
    img = Image.new("RGB", (80, 80), (255, 220, 120))
    d = ImageDraw.Draw(img)
    d.ellipse([10, 10, 70, 70], fill="orange", outline="red", width=3)
    d.text((24, 30), "?!?", fill="black", font=_get_font(22))
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
            run.font.name = font_name
            run.font.bold = bold
    return box


def set_flip_vertical(pic) -> None:
    """Повернуть картинку вертикально (flipV=1) — атрибут прямо в XML слайда."""
    xfrm = pic._element.spPr.find(qn("a:xfrm"))
    if xfrm is None:
        raise RuntimeError("у картинки нет xfrm — сначала задай позицию")
    xfrm.set("flipV", "1")


def set_auto_advance(slide, ms: int) -> None:
    """Включить автопереключение слайда (слайды начнут «ездить сами»)."""
    sld = slide._element
    transition = sld.makeelement(
        qn("p:transition"), {"advClick": "0", "advTm": str(ms)}
    )
    transition.append(transition.makeelement(qn("p:fade"), {}))
    clr = sld.find(qn("p:clrMapOvr"))
    anchor = clr if clr is not None else sld.find(qn("p:cSld"))
    anchor.addnext(transition)  # порядок в XML: cSld -> clrMapOvr -> transition


def new_presentation() -> Presentation:
    prs = Presentation()
    prs.slide_width = Inches(13.333)   # 16:9, как у современных проекторов
    prs.slide_height = Inches(7.5)
    return prs


# ---------------------------------------------------------------- файлы

def build_broken(path: Path, scheme: Path, tiny: Path) -> list:
    prs = new_presentation()
    blank = prs.slide_layouts[6]
    expected = []

    # --- слайд 1: обложка (дефектов нет) ---
    s = prs.slides.add_slide(blank)
    add_text(s, 1.5, 2.6, 10.5, 1.6,
             "Тестовая презентация «СЛОМАННАЯ»\n"
             "Каждый слайд содержит ровно один дефект", 32, bold=True)
    add_text(s, 1.5, 4.6, 10.5, 1.0,
             "Эталон для проверки программы (см. expected_defects.json)", 20)

    # --- слайд 2: перевёрнутая картинка ---
    s = prs.slides.add_slide(blank)
    add_text(s, 0.5, 0.3, 12, 0.9, "Слайд 2. Схема процесса", 28, bold=True)
    pic = s.shapes.add_picture(str(scheme), Inches(3.5), Inches(1.6),
                               Inches(6), Inches(4))
    set_flip_vertical(pic)                       # ← ДЕФЕКТ 1
    expected.append({"slide": 2, "defect": "flipped_image",
                     "description": "картинка перевёрнута вертикально (flipV=1)"})

    # --- слайд 3: объекты вылезают за границы слайда ---
    s = prs.slides.add_slide(blank)
    add_text(s, 0.5, 0.3, 12, 0.9, "Слайд 3. Длинный текст", 28, bold=True)
    add_text(s, 0.7, 5.6, 8, 3.2,
             "Этот текстовый блок начинается слишком низко и заканчивается "
             "значительно ниже нижней границы слайда. В режиме показа последние "
             "строки просто не будут видны. Классическая ошибка: текст не "
             "поместился, а вы его не проверили перед уроком.", 22)
    expected.append({"slide": 3, "defect": "object_out_of_slide",
                     "description": "текстовое поле: низ (5.6+3.2=8.8\") > высоты слайда (7.5\")"})
    pic = s.shapes.add_picture(str(scheme), Inches(10.5), Inches(1.6),
                               Inches(4), Inches(2.7))   # право = 14.5" > 13.333"
    expected.append({"slide": 3, "defect": "object_out_of_slide",
                     "description": "картинка: правый край (10.5+4=14.5\") > ширины слайда (13.333\")"})

    # --- слайд 4: картинка перекрывает текст ---
    s = prs.slides.add_slide(blank)
    add_text(s, 0.5, 0.3, 12, 0.9, "Слайд 4. Важные определения", 28, bold=True)
    add_text(s, 0.8, 1.6, 6, 3.4,
             "Здесь должен быть текст определения. Он хорошо читался, пока "
             "поверх него не вставили картинку. Правая часть абзацев "
             "оказалась под изображением — на проекторе это выглядит как "
             "мусор.", 20)
    s.shapes.add_picture(str(scheme), Inches(4.5), Inches(2.6),
                         Inches(6), Inches(4))           # пересекается с текстом
    expected.append({"slide": 4, "defect": "overlap_picture_text",
                     "description": "прямоугольники текста (0.8..6.8) и картинки (4.5..10.5) пересекаются"})

    # --- слайд 5: картинка низкого разрешения ---
    s = prs.slides.add_slide(blank)
    add_text(s, 0.5, 0.3, 12, 0.9, "Слайд 5. Иллюстрация", 28, bold=True)
    add_text(s, 0.7, 1.4, 6, 2.5,
             "Картинка 80x80 пикселей растянута на 6 дюймов. "
             "Эффективное разрешение ≈ 13 DPI — на проекторе будут "
             "огромные пиксели.", 20)
    s.shapes.add_picture(str(tiny), Inches(7), Inches(1.6),
                         Inches(6), Inches(5.5))         # 80px на 6" = 13 DPI
    expected.append({"slide": 5, "defect": "low_resolution_image",
                     "description": "80x80 px при отображении 6\"x5.5\" → ~13 DPI (< 150)"})

    # --- слайд 6: нестандартный шрифт ---
    s = prs.slides.add_slide(blank)
    add_text(s, 0.5, 0.3, 12, 0.9, "Слайд 6. Нестандартный шрифт", 28, bold=True)
    add_text(s, 0.7, 1.6, 11, 3.0,
             "Этот текст набран шрифтом Comic Sans MS, который не входит в "
             "стандартный набор и не внедрён в файл. На компьютере в классе "
             "его может не быть — и PowerPoint молча подменит шрифт, "
             "а вёрстка поедет.", 22, font_name="Comic Sans MS")
    expected.append({"slide": 6, "defect": "nonstandard_font",
                     "description": "используется шрифт 'Comic Sans MS' (не из белого списка, не внедрён)"})

    # --- слайд 7: автопереход ---
    s = prs.slides.add_slide(blank)
    add_text(s, 0.5, 0.3, 12, 0.9, "Слайд 7. Автопереход", 28, bold=True)
    add_text(s, 0.7, 1.6, 11, 2.5,
             "У этого слайда забыт автопереключение через 2 секунды. "
             "На уроке слайды начнут «ездить сами» — учитель не успевает "
             "говорить.", 22)
    set_auto_advance(s, 2000)                    # ← ДЕФЕКТ
    expected.append({"slide": 7, "defect": "auto_advance_transition",
                     "description": "<p:transition advTm=\"2000\" advClick=\"0\"> — сам переключится через 2 с"})

    # --- слайд 8: внешняя ссылка ---
    s = prs.slides.add_slide(blank)
    add_text(s, 0.5, 0.3, 12, 0.9, "Слайд 8. Внешняя ссылка", 28, bold=True)
    add_text(s, 0.7, 1.6, 11, 2.0,
             "Слайд ссылается на видео во внешнем файле/интернете. "
             "В школе без интернета или на чужом компьютере ссылка не "
             "откроется.", 22)
    box = add_text(s, 0.7, 4.0, 6, 1.0, "Открыть видео", 24)
    run = box.text_frame.paragraphs[0].runs[0]
    run.hyperlink.address = "https://example.com/video-lesson"   # ← ДЕФЕКТ
    expected.append({"slide": 8, "defect": "external_link",
                     "description": "ран с внешним hyperlink (TargetMode=External)"})

    prs.save(path)
    return expected


def build_clean(path: Path, scheme: Path) -> None:
    prs = new_presentation()
    blank = prs.slide_layouts[6]

    s = prs.slides.add_slide(blank)
    add_text(s, 1.5, 2.8, 10.5, 1.5,
             "Тестовая презентация «ЧИСТАЯ»\nЗдоровый файл без дефектов", 32, bold=True)

    s = prs.slides.add_slide(blank)
    add_text(s, 0.5, 0.3, 12, 0.9, "Всё в порядке", 28, bold=True)
    add_text(s, 0.8, 1.6, 5.5, 3.5,
             "Стандартный шрифт Arial, текст внутри границ слайда, "
             "картинка не перекрывает текст и не вылезает за край. "
             "Такой файл программа должна пропустить без единого "
             "предупреждения.", 20)
    s.shapes.add_picture(str(scheme), Inches(7), Inches(1.6),
                         Inches(5.5), Inches(3.7))

    s = prs.slides.add_slide(blank)
    add_text(s, 0.5, 0.3, 12, 0.9, "Второй слайд — тоже без ошибок", 28, bold=True)
    add_text(s, 0.8, 1.6, 11, 4.0,
             "Простой текст. Если программа найдёт здесь проблему — "
             "это ложное срабатывание, и её нужно дорабатывать.", 22)

    prs.save(path)


def main() -> None:
    OUT.mkdir(exist_ok=True)
    ASSETS.mkdir(exist_ok=True)

    scheme = ASSETS / "scheme.png"
    tiny = ASSETS / "tiny.png"
    make_scheme(scheme)
    make_tiny(tiny)

    broken = OUT / "broken.pptx"
    clean = OUT / "clean.pptx"
    defects = build_broken(broken, scheme, tiny)
    build_clean(clean, scheme)

    expected = {
        "broken.pptx": defects,
        "clean.pptx": [],
    }
    (OUT / "expected_defects.json").write_text(
        json.dumps(expected, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    print("Готово! Созданы файлы:")
    for f in sorted(OUT.rglob("*")):
        if f.is_file():
            print(f"  {f.relative_to(BASE)}  ({f.stat().st_size} байт)")
    print(f"\nДефектов в эталоне: {len(defects)} (на 7 слайдах из 8)")


if __name__ == "__main__":
    main()
