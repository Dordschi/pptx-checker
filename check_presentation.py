#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_presentation.py — проверка презентации .pptx перед показом на уроке.

Чек-лист (7 проверок):
  1. flipped_image           — перевёрнутые картинки (flipV / flipH в XML)
  2. object_out_of_slide     — объекты, вылезающие за границы слайда
  3. overlap_picture_text    — картинка перекрывает текст
  4. low_resolution_image    — картинка низкого разрешения (эффективный DPI < 150)
  5. nonstandard_font        — нестандартный шрифт (может слететь на чужом ПК)
  6. auto_advance_transition — забытый автопереход (слайды едут сами)
  7. external_link           — внешняя ссылка (не откроется без интернета)

Использование:
  python3 check_presentation.py файл.pptx         — отчёт о проблемах
  python3 check_presentation.py файл.pptx --fix   — исправить безопасные
                                                     дефекты в КОПИЮ файла
  python3 check_presentation.py --test            — самопроверка на эталоне
"""

import argparse
import json
import sys
from collections import Counter
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path

from PIL import Image as PILImage
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

# Windows: не падать на кириллице в консоли (cp1252/cp866)
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

BASE = Path(__file__).resolve().parent
TEST_FILES = BASE / "test_files"

IMPLEMENTED = {
    "flipped_image",
    "object_out_of_slide",
    "overlap_picture_text",
    "low_resolution_image",
    "nonstandard_font",
    "auto_advance_transition",
    "external_link",
}

EMU_PER_INCH = 914400
MIN_DPI = 150

# Шрифты, которые есть в Windows/Office «из коробки» — на любом школьном ПК
STANDARD_FONTS = {
    "arial", "calibri", "times new roman", "verdana", "tahoma",
    "georgia", "cambria", "trebuchet ms", "courier new",
}


# ------------------------------------------------------------------ модель

@dataclass
class Issue:
    slide: int       # номер слайда (с 1)
    defect: str      # id дефекта (как в expected_defects.json)
    message: str     # читаемое описание для отчёта


def _in(emu: int) -> str:
    """EMU → дюймы с двумя знаками (1 дюйм = 914400 EMU)."""
    return f'{emu / EMU_PER_INCH:.2f}"'


def _shape_label(shape) -> str:
    kinds = {
        MSO_SHAPE_TYPE.PICTURE: "Рисунок",
        MSO_SHAPE_TYPE.TEXT_BOX: "Текстовое поле",
        MSO_SHAPE_TYPE.AUTO_SHAPE: "Фигура",
    }
    kind = kinds.get(shape.shape_type, "Объект")
    return f"«{shape.name}» ({kind})"


def _rect(shape):
    """Прямоугольник объекта: (лево, верх, право, низ) в EMU."""
    return (shape.left, shape.top,
            shape.left + shape.width, shape.top + shape.height)


def _rects_overlap(a, b) -> bool:
    """Есть ли непустая общая область у двух прямоугольников?"""
    return not (a[2] <= b[0] or b[2] <= a[0] or
                a[3] <= b[1] or b[3] <= a[1])


# ------------------------------------------------------------------ проверки

def check_flipped_images(slide, slide_no: int, issues: list) -> None:
    """1. Картинка перевёрнута вертикально/горизонтально?"""
    for shape in slide.shapes:
        if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
            continue
        xfrm = shape._element.spPr.find(qn("a:xfrm"))
        if xfrm is None:
            continue
        if xfrm.get("flipV") == "1":
            issues.append(Issue(
                slide_no, "flipped_image",
                f"{_shape_label(shape)} перевёрнута вертикально (flipV)"))
        if xfrm.get("flipH") == "1":
            issues.append(Issue(
                slide_no, "flipped_image",
                f"{_shape_label(shape)} перевёрнута горизонтально (flipH)"))


def check_bounds(slide, slide_no: int, slide_w: int, slide_h: int,
                 issues: list) -> None:
    """2. Выходит ли объект за границы слайда?"""
    for shape in slide.shapes:
        left, top = shape.left, shape.top
        width, height = shape.width, shape.height
        if None in (left, top, width, height):
            continue
        right, bottom = left + width, top + height
        over = []
        if left < 0:
            over.append(f"левый край {_in(left)} < 0\"")
        if top < 0:
            over.append(f"верх {_in(top)} < 0\"")
        if right > slide_w:
            over.append(f"правый край {_in(right)} > {_in(slide_w)}")
        if bottom > slide_h:
            over.append(f"низ {_in(bottom)} > {_in(slide_h)}")
        if over:
            issues.append(Issue(
                slide_no, "object_out_of_slide",
                f"{_shape_label(shape)} вылезает за слайд: " + "; ".join(over)))


def check_overlap(slide, slide_no: int, issues: list) -> None:
    """3. Картинка перекрывает текстовый блок?"""
    pictures = [s for s in slide.shapes
                if s.shape_type == MSO_SHAPE_TYPE.PICTURE
                and None not in (s.left, s.top, s.width, s.height)]
    text_shapes = [s for s in slide.shapes
                   if getattr(s, "has_text_frame", False)
                   and s.text_frame.text.strip()
                   and None not in (s.left, s.top, s.width, s.height)]
    for pic in pictures:
        pr = _rect(pic)
        for txt in text_shapes:
            if _rects_overlap(pr, _rect(txt)):
                issues.append(Issue(
                    slide_no, "overlap_picture_text",
                    f"{_shape_label(pic)} перекрывает "
                    f"{_shape_label(txt)}"))
                break  # одна жалоба на картинку — не задваиваем


def check_low_resolution(slide, slide_no: int, issues: list) -> None:
    """4. Картинка растянута с низким разрешением (DPI < 150)?"""
    for shape in slide.shapes:
        if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
            continue
        if None in (shape.width, shape.height) or shape.width <= 0:
            continue
        try:
            im = PILImage.open(BytesIO(shape.image.blob))
            px_w, px_h = im.size
        except Exception:
            continue
        w_in = shape.width / EMU_PER_INCH
        h_in = shape.height / EMU_PER_INCH
        dpi = min(px_w / w_in, px_h / h_in)
        if dpi < MIN_DPI:
            issues.append(Issue(
                slide_no, "low_resolution_image",
                f"{_shape_label(shape)} ~{dpi:.0f} DPI "
                f"({px_w}x{px_h} px на {w_in:.1f}\"x{h_in:.1f}\") — "
                f"на проекторе будут пиксели"))


def check_fonts(slide, slide_no: int, issues: list) -> None:
    """5. Нестандартный шрифт, который может слететь на чужом ПК?"""
    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        bad = set()
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                name = run.font.name
                if name and name.lower() not in STANDARD_FONTS:
                    bad.add(name)
        if bad:
            issues.append(Issue(
                slide_no, "nonstandard_font",
                f"{_shape_label(shape)}: шрифт {', '.join(sorted(bad))} — "
                f"на школьном компьютере может замениться, вёрстка поедет"))


def check_auto_transition(slide, slide_no: int, issues: list) -> None:
    """6. Забытый автопереход — слайд переключится сам?"""
    tr = slide._element.find(qn("p:transition"))
    if tr is not None and tr.get("advTm"):
        sec = int(tr.get("advTm")) / 1000
        issues.append(Issue(
            slide_no, "auto_advance_transition",
            f"автопереключение через {sec:g} с — слайды поедут сами"))


def check_external_links(slide, slide_no: int, issues: list) -> None:
    """7. Внешние ссылки/ресурсы — не откроются офлайн?"""
    rels = slide.part.rels
    try:
        items = list(rels.values())
    except Exception:
        items = list(rels)
    for rel in items:
        if getattr(rel, "is_external", False):
            issues.append(Issue(
                slide_no, "external_link",
                f"внешняя ссылка: {getattr(rel, 'target_ref', '?')} — "
                f"на школьном ПК без интернета не откроется"))


# ------------------------------------------------------------------ автопочинка (--fix)

def fix_presentation(src_path, dst_path=None):
    """
    Исправляет БЕЗОПАСНЫЕ дефекты и сохраняет результат копией.
    Исходный файл не изменяется. Возвращает (fixed, dst_path).

    Чиним только то, где признак объективен и правка ничего не ломает:
      - flipped_image: снимаем атрибуты flipV/flipH (просто удаляем их)
      - auto_advance_transition: убираем advTm, включаем переход по клику
    Остальные дефекты (вылезание, перекрытие, шрифты...) остаются
    предупреждениями: двигать объекты наугад опасно для вёрстки.
    """
    prs = Presentation(src_path)
    fixed = []

    for slide_no, slide in enumerate(prs.slides, 1):
        # --- 1) перевёрнутые картинки: снимаем отражение ---
        for shape in slide.shapes:
            if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
                continue
            xfrm = shape._element.spPr.find(qn("a:xfrm"))
            if xfrm is None:
                continue
            for attr in ("flipV", "flipH"):
                if xfrm.get(attr) == "1":
                    del xfrm.attrib[attr]      # убрали признак переворота
                    fixed.append(Issue(
                        slide_no, "flipped_image",
                        f"{_shape_label(shape)}: снято отражение {attr}"))

        # --- 2) автопереход: убираем advTm, оставляем переход по клику ---
        tr = slide._element.find(qn("p:transition"))
        if tr is not None and tr.get("advTm"):
            del tr.attrib["advTm"]             # больше сам не переключится
            tr.set("advClick", "1")            # но кликнуть можно
            fixed.append(Issue(
                slide_no, "auto_advance_transition",
                "автопереход отключён — слайд переключается по клику"))

    # сохраняем всегда КОПИЕЙ — исходник не трогаем
    if dst_path is None:
        p = Path(src_path)
        dst_path = p.with_name(p.stem + "_исправлено" + p.suffix)
    prs.save(dst_path)
    return fixed, Path(dst_path)


# ------------------------------------------------------------------ главная функция

def check_presentation(path) -> list:
    """Прогоняет все 7 проверок. Возвращает список Issue."""
    prs = Presentation(path)
    issues: list = []
    for slide_no, slide in enumerate(prs.slides, 1):
        check_flipped_images(slide, slide_no, issues)
        check_bounds(slide, slide_no, prs.slide_width, prs.slide_height,
                     issues)
        check_overlap(slide, slide_no, issues)
        check_low_resolution(slide, slide_no, issues)
        check_fonts(slide, slide_no, issues)
        check_auto_transition(slide, slide_no, issues)
        check_external_links(slide, slide_no, issues)
    return issues


# ------------------------------------------------------------------ отчёт

def print_report(path, issues) -> None:
    prs = Presentation(path)
    n_slides = len(prs.slides._sldIdLst)
    print(f"ПРОВЕРКА: {path}")
    print(f"Слайдов: {n_slides}. Проверок: 7 (перевёрнутые картинки, "
          f"вылезание за границы, перекрытие, разрешение, шрифты, "
          f"автопереход, ссылки).")
    print("-" * 60)
    if not issues:
        print("✅ Проблем не найдено.")
    else:
        by_slide = {}
        for issue in issues:
            by_slide.setdefault(issue.slide, []).append(issue)
        for slide_no in sorted(by_slide):
            # одинаковые строки на одном слайде склеиваем: «… ×N»
            counts = {}
            order = []
            for issue in by_slide[slide_no]:
                key = (issue.defect, issue.message)
                if key in counts:
                    counts[key] += 1
                else:
                    counts[key] = 1
                    order.append(key)
            for defect, message in order:
                n = counts[(defect, message)]
                tail = f"  ×{n}" if n > 1 else ""
                print(f"⚠️  Слайд {slide_no}: [{defect}] {message}{tail}")
    print("-" * 60)
    print(f"ИТОГО: проблем — {len(issues)}")
    if issues:
        print("Рекомендация: исправьте найденное или сохраните копию в PDF "
              "как страховку.")


# ------------------------------------------------------------------ самопроверка (--test)

def run_test() -> int:
    expected_path = TEST_FILES / "expected_defects.json"
    if not expected_path.exists():
        print("Не найден test_files/expected_defects.json. "
              "Сначала запусти: python3 generate_test_files.py")
        return 2
    expected = json.loads(expected_path.read_text(encoding="utf-8"))

    all_ok = True
    for fname, expected_list in expected.items():
        path = TEST_FILES / fname
        print(f"\n=== {fname} ===")
        found = check_presentation(path)

        exp_counter = Counter(
            (e["slide"], e["defect"]) for e in expected_list
            if e["defect"] in IMPLEMENTED
        )
        found_counter = Counter((i.slide, i.defect) for i in found)

        for key, count in sorted(exp_counter.items()):
            got = found_counter.get(key, 0)
            mark = "✅" if got == count else "❌"
            if got != count:
                all_ok = False
            print(f"  {mark} слайд {key[0]}: {key[1]} — найдено {got} из {count}")

        extra = found_counter - exp_counter
        for key, count in sorted(extra.items()):
            all_ok = False
            print(f"  ❌ ЛОЖНОЕ СРАБАТЫВАНИЕ: слайд {key[0]}: {key[1]} ×{count}")

        not_impl = {(e["slide"], e["defect"]) for e in expected_list
                    if e["defect"] not in IMPLEMENTED}
        for slide_no, defect in sorted(not_impl):
            print(f"  ⏳ слайд {slide_no}: {defect} — проверка ещё не реализована")

        if not expected_list and not found:
            print("  ✅ чистый файл принят (ложных срабатываний нет)")

    print("\n" + "=" * 60)
    print("ИТОГ ТЕСТА:", "PASS ✅" if all_ok else "FAIL ❌")
    return 0 if all_ok else 1


# ------------------------------------------------------------------ CLI

def main() -> int:
    parser = argparse.ArgumentParser(
        description="Проверка .pptx перед уроком (7 проверок)")
    parser.add_argument("file", nargs="?", help="путь к презентации .pptx")
    parser.add_argument("--test", action="store_true",
                        help="самопроверка на эталонных файлах")
    parser.add_argument("--fix", action="store_true",
                        help="исправить безопасные дефекты в копию файла")
    args = parser.parse_args()

    if args.test:
        return run_test()
    if not args.file:
        # Двойной клик по программе (без аргументов) — открыть окно с интерфейсом
        try:
            import app_ui
            print("Окно программы открыто. Закройте его, чтобы завершить.")
            app_ui.root.mainloop()
            return 0
        except Exception:
            # tkinter недоступен (например, Linux без python3-tk) — fallback
            parser.print_help()
            print("\nОкно не открылось. Запуск из командной строки: см. README.")
            if sys.stdin is not None and sys.stdin.isatty():
                try:
                    input("Нажмите Enter, чтобы закрыть окно...")
                except EOFError:
                    pass
            return 2
    path = Path(args.file)
    if not path.exists():
        print(f"Файл не найден: {path}")
        return 2

    if args.fix:
        try:
            before = check_presentation(path)
            fixed, dst = fix_presentation(path)
            after = check_presentation(dst)
        except Exception:
            print(f"❌ Не удалось открыть файл: {path}")
            print("   Архив .pptx повреждён или это не презентация.")
            return 3
        print(f"ИСХОДНИК НЕ ИЗМЕНЁН: {path}")
        print(f"Создана копия:       {dst}")
        print(f"Исправлено: {len(fixed)}")
        for f in fixed:
            print(f"  ✔ слайд {f.slide}: {f.message}")
        print(f"Проблем до: {len(before)} → после: {len(after)}")
        print()
        print_report(dst, after)
        return 0

    try:
        issues = check_presentation(path)
    except Exception:
        print(f"❌ Не удалось открыть файл: {path}")
        print("   Архив .pptx повреждён или это не презентация.")
        print("   (Такой файл не проверить — скачайте презентацию заново.)")
        return 3
    print_report(path, issues)
    return 1 if issues else 0


if __name__ == "__main__":
    sys.exit(main())
