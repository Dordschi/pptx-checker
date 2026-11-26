#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_presentation.py — проверка учебной презентации .pptx перед уроком.

Проверки:
  1. corrupted_media         — битые картинки/медиа внутри архива .pptx
  2. flipped_image           — перевёрнутые вверх ногами картинки (flipV)
  3. object_out_of_slide     — видимые объекты выходят за границы слайда
  4. overlap_picture_text    — непрозрачная картинка сверху закрывает текст
  5. low_resolution_image    — растянутая картинка низкого качества (пиксели на проекторе)
  6. nonstandard_font        — нестандартный и не встроенный в файл шрифт
  7. small_font              — слишком мелкий текст (< 14 pt), нечитаемый с задних парт
  8. auto_advance_transition — забытый автопереход по таймеру / отключённый клик
  9. external_link           — внешние веб-ссылки и невнедрённые локальные файлы

Использование:
  python3 check_presentation.py презентация.pptx       — отчёт о проблемах
  python3 check_presentation.py папка_с_уроками/       — сводная таблица по папке
  python3 check_presentation.py презентация.pptx --fix — исправить безопасные
                                                         дефекты в КОПИЮ файла
  python3 check_presentation.py --test                 — самопроверка на эталоне
"""

from __future__ import annotations

import argparse
import json
import math
import sys
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from typing import Iterable, List, Optional, Sequence, Tuple

from lxml import etree
from PIL import Image as PILImage
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.exc import PackageNotFoundError
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

EMU_PER_INCH = 914400
EMU_PER_CM = 360000          # 1 дюйм = 2.54 см = 914400 EMU -> 1 см = 360000 EMU
EMU_PER_PT = 12700           # 1 пункт = 1/72 дюйма = 12700 EMU

# Пороги проверок, подобранные под реальные школьные проекторы и шаблоны PowerPoint:
OUT_OF_SLIDE_TOLERANCE = int(0.3 * EMU_PER_CM)   # 3 мм — допуск на округление координат
OVERLAP_MIN_DIM = int(0.3 * EMU_PER_CM)          # минимум 3 мм пересечения по X и Y
MIN_DPI_MAIN = 75            # ниже 75 DPI на проекторе заметны пиксели («мыло»)
MIN_DPI_ICON = 50            # для мелкого декора (< 3.8 см) порог мягче
MIN_FONT_PT = 14.0           # мельче 14 pt с задних парт в классе не прочитать

IMPLEMENTED = {
    "corrupted_media",
    "flipped_image",
    "object_out_of_slide",
    "overlap_picture_text",
    "low_resolution_image",
    "nonstandard_font",
    "small_font",
    "auto_advance_transition",
    "external_link",
}

DEFECT_TITLES = {
    "corrupted_media": "Повреждённый медиафайл",
    "flipped_image": "Перевёрнутая картинка",
    "object_out_of_slide": "Выход за границы слайда",
    "overlap_picture_text": "Картинка перекрывает текст",
    "low_resolution_image": "Низкое разрешение картинки",
    "nonstandard_font": "Нестандартный шрифт",
    "small_font": "Слишком мелкий шрифт",
    "auto_advance_transition": "Автопереключение слайда",
    "external_link": "Внешняя ссылка / файл",
}

# Дефекты, которые можно гарантированно безопасно исправить флагом --fix
FIXABLE_DEFECTS = {"flipped_image", "auto_advance_transition"}

# Стандартные шрифты Windows и MS Office, установленные на школьных ПК «из коробки»
STANDARD_FONTS = {
    "arial", "arial black", "arial narrow",
    "calibri", "calibri light",
    "times new roman",
    "verdana", "tahoma", "georgia",
    "cambria", "cambria math",
    "trebuchet ms", "courier new", "consolas", "lucida console", "lucida sans unicode",
    "segoe ui", "segoe ui light", "segoe ui semibold", "segoe print",
    "candara", "corbel", "constantia",
    "comic sans ms", "impact", "franklin gothic medium",
    "palatino linotype", "book antiqua", "garamond", "century gothic", " century",
    "microsoft sans serif", "ms sans serif", "ms serif",
    "symbol", "wingdings", "wingdings 2", "wingdings 3", "webdings",
    # Свободные метрические аналоги (LibreOffice / российские школьные ОС)
    "liberation sans", "liberation serif", "liberation mono",
    "dejavu sans", "dejavu serif", "dejavu sans mono",
}

# Типы заполнителей (колонтитулы, дата, номер слайда), где мелкий кегль — это норма
FOOTER_PLACEHOLDERS = {
    PP_PLACEHOLDER.FOOTER,
    PP_PLACEHOLDER.DATE,
    PP_PLACEHOLDER.SLIDE_NUMBER,
    PP_PLACEHOLDER.HEADER,
}

OLE2_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


# ------------------------------------------------------------------ модели и ошибки

class PresentationOpenError(Exception):
    """Ошибка открытия файла презентации (не .pptx, старый .ppt или битый архив)."""


@dataclass
class Issue:
    slide: int                   # номер слайда (с 1; 0 — общая проблема файла)
    defect: str                  # идентификатор проверки (см. IMPLEMENTED)
    message: str                 # понятное учителю описание проблемы
    severity: str = "warning"    # "error" (критично) | "warning" (внимание)


@dataclass
class InspectionResult:
    path: Path
    slide_count: int
    slide_width_cm: float
    slide_height_cm: float
    issues: List[Issue] = field(default_factory=list)

    @property
    def error_count(self) -> int:
        return sum(1 for i in self.issues if i.severity == "error")

    @property
    def warning_count(self) -> int:
        return sum(1 for i in self.issues if i.severity != "error")

    @property
    def fixable_count(self) -> int:
        return sum(1 for i in self.issues if i.defect in FIXABLE_DEFECTS)


@dataclass
class ShapeNode:
    """Объект слайда с реальными координатами на листе и порядком слоя (z-order)."""
    shape: object
    z_order: Tuple[int, ...]
    left: float
    top: float
    width: float
    height: float
    rotation_deg: float = 0.0
    flip_h: bool = False
    flip_v: bool = False
    in_group: bool = False


# ------------------------------------------------------------------ вспомогательные функции

def _cm(emu: float) -> str:
    """Перевод EMU в сантиметры с одним знаком после запятой."""
    val = emu / EMU_PER_CM
    return f"{val:.1f} см".replace(".", ",")


def _clean_snippet(text: str, limit: int = 28) -> str:
    """Короткий читаемый фрагмент текста объекта для подписи в отчёте."""
    cleaned = " ".join(text.replace("\x0b", " ").split())
    if not cleaned:
        return ""
    if len(cleaned) > limit:
        return cleaned[: limit - 1].rstrip() + "…"
    return cleaned


def _extract_shape_text(shape) -> str:
    """Достаёт видимый текст из текстового поля, фигуры или таблицы."""
    if getattr(shape, "has_text_frame", False):
        return shape.text_frame.text or ""
    if getattr(shape, "has_table", False):
        parts = []
        for row in shape.table.rows:
            for cell in row.cells:
                t = cell.text_frame.text.strip()
                if t:
                    parts.append(t)
        return " | ".join(parts)
    return ""


def _shape_label(shape) -> str:
    """Понятное учителю название объекта на слайде (с фрагментом текста, если он есть)."""
    kinds = {
        MSO_SHAPE_TYPE.PICTURE: "Рисунок",
        MSO_SHAPE_TYPE.TEXT_BOX: "Текстовое поле",
        MSO_SHAPE_TYPE.AUTO_SHAPE: "Фигура",
        MSO_SHAPE_TYPE.PLACEHOLDER: "Блок макета",
        MSO_SHAPE_TYPE.TABLE: "Таблица",
        MSO_SHAPE_TYPE.GROUP: "Группа объектов",
        MSO_SHAPE_TYPE.CHART: "Диаграмма",
        MSO_SHAPE_TYPE.MEDIA: "Медиафайл",
        MSO_SHAPE_TYPE.LINE: "Линия",
    }
    kind = kinds.get(shape.shape_type, "Объект")
    if shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
        try:
            ph_type = shape.placeholder_format.type
            if ph_type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
                kind = "Заголовок"
            elif ph_type == PP_PLACEHOLDER.SUBTITLE:
                kind = "Подзаголовок"
            elif ph_type == PP_PLACEHOLDER.PICTURE:
                kind = "Рисунок"
        except Exception:
            pass

    snippet = _clean_snippet(_extract_shape_text(shape))
    if snippet:
        return f"{kind} «{snippet}»"
    return f"{kind} «{shape.name}»"


def _read_xfrm_attrs(shape) -> Tuple[float, bool, bool]:
    """Читает угол поворота (в градусах) и флаги flipH / flipV из XML объекта."""
    el = shape._element
    xfrm = None
    for container_tag in ("p:spPr", "p:grpSpPr"):
        container = el.find(qn(container_tag))
        if container is not None:
            xfrm = container.find(qn("a:xfrm"))
            if xfrm is not None:
                break
    if xfrm is None:
        xfrm = el.find(qn("p:xfrm"))
    if xfrm is None:
        xfrm = el.find(qn("a:xfrm"))
    if xfrm is None:
        return 0.0, False, False

    try:
        rot_deg = (int(xfrm.get("rot", "0")) / 60000.0) % 360.0
    except ValueError:
        rot_deg = 0.0
    flip_h = xfrm.get("flipH") in ("1", "true")
    flip_v = xfrm.get("flipV") in ("1", "true")
    return rot_deg, flip_h, flip_v


def _iter_slide_nodes(shapes, prefix: Tuple[int, ...] = (),
                      transform=None) -> Iterable[ShapeNode]:
    """
    Рекурсивно обходит все объекты слайда (включая вложенные в группы Ctrl+G),
    сохраняя точный порядок наложения слоёв z_order и пересчитывая координаты.
    """
    for idx, shape in enumerate(shapes):
        z_cur = prefix + (idx,)
        left, top, width, height = shape.left, shape.top, shape.width, shape.height
        if None in (left, top, width, height):
            continue

        rot_deg, flip_h, flip_v = _read_xfrm_attrs(shape)
        if transform is not None:
            grp_x, grp_y, ch_x, ch_y, sx, sy, g_rot, g_fh, g_fv, ch_w, ch_h = transform
            rel_x = (left - ch_x) * sx
            rel_y = (top - ch_y) * sy
            real_w = width * sx
            real_h = height * sy
            if g_fh:
                grp_w = sx * ch_w
                rel_x = grp_w - rel_x - real_w
            if g_fv:
                grp_h = sy * ch_h
                rel_y = grp_h - rel_y - real_h
            real_x = grp_x + rel_x
            real_y = grp_y + rel_y
            rot_deg = (g_rot + rot_deg) % 360.0
            flip_h = g_fh ^ flip_h
            flip_v = g_fv ^ flip_v
            in_grp = True
        else:
            real_x, real_y, real_w, real_h = float(left), float(top), float(width), float(height)
            in_grp = False

        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            grp_sp_pr = shape._element.find(qn("p:grpSpPr"))
            xfrm = grp_sp_pr.find(qn("a:xfrm")) if grp_sp_pr is not None else None
            ch_off = xfrm.find(qn("a:chOff")) if xfrm is not None else None
            ch_ext = xfrm.find(qn("a:chExt")) if xfrm is not None else None
            ch_x = ch_y = ch_w = ch_h = 0
            if ch_off is not None and ch_ext is not None:
                try:
                    ch_x = int(ch_off.get("x", "0"))
                    ch_y = int(ch_off.get("y", "0"))
                    ch_w = int(ch_ext.get("cx", "0"))
                    ch_h = int(ch_ext.get("cy", "0"))
                except ValueError:
                    ch_w = ch_h = 0
            if ch_w > 0 and ch_h > 0:
                sx = real_w / ch_w
                sy = real_h / ch_h
                child_tf = (real_x, real_y, ch_x, ch_y, sx, sy,
                            rot_deg, flip_h, flip_v, ch_w, ch_h)
            else:
                child_tf = (0.0, 0.0, 0, 0, 1.0, 1.0,
                            rot_deg, flip_h, flip_v, int(real_w), int(real_h))
            yield from _iter_slide_nodes(shape.shapes, z_cur, child_tf)
            continue

        yield ShapeNode(
            shape=shape,
            z_order=z_cur,
            left=real_x,
            top=real_y,
            width=real_w,
            height=real_h,
            rotation_deg=rot_deg,
            flip_h=flip_h,
            flip_v=flip_v,
            in_group=in_grp,
        )


def _aabb(node: ShapeNode) -> Tuple[float, float, float, float]:
    """Истинный ограничивающий прямоугольник (лево, верх, право, низ) с учётом поворота."""
    if node.rotation_deg % 360.0 == 0.0:
        return (node.left, node.top, node.left + node.width, node.top + node.height)
    cx = node.left + node.width / 2.0
    cy = node.top + node.height / 2.0
    rad = math.radians(node.rotation_deg)
    cos_a = abs(math.cos(rad))
    sin_a = abs(math.sin(rad))
    eff_w = node.width * cos_a + node.height * sin_a
    eff_h = node.width * sin_a + node.height * cos_a
    return (cx - eff_w / 2.0, cy - eff_h / 2.0, cx + eff_w / 2.0, cy + eff_h / 2.0)


def _get_picture_blob(shape) -> Optional[bytes]:
    """Возвращает байты изображения, если объект — картинка или заполнитель с рисунком."""
    if shape.shape_type not in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.PLACEHOLDER):
        return None
    if not hasattr(shape, "image"):
        return None
    try:
        blob = shape.image.blob
        return blob if blob else None
    except Exception:
        return None


def _has_visible_box_styling(shape) -> bool:
    """Есть ли у пустой фигуры видимая заливка или рамка?"""
    sp_pr = getattr(shape._element, "spPr", None)
    if sp_pr is None:
        return False
    for fill_tag in ("a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill"):
        if sp_pr.find(qn(fill_tag)) is not None:
            return True
    ln = sp_pr.find(qn("a:ln"))
    if ln is not None and ln.find(qn("a:noFill")) is None:
        if ln.find(qn("a:solidFill")) is not None or ln.find(qn("a:gradFill")) is not None:
            return True
    return False


# ------------------------------------------------------------------ безопасное открытие .pptx

def _open_presentation_safely(path: Path) -> Tuple[Presentation, List[str]]:
    """
    Открывает .pptx и аккуратно обрабатывает:
      - старый бинарный формат .ppt (OLE2 PowerPoint 97–2003);
      - полностью битые архивы / чужие файлы;
      - частично битые архивы, где повреждён только медиафайл в ppt/media/*
        (позволяет проверить все слайды и точно указать битую картинку).
    """
    if not path.exists() or not path.is_file():
        raise FileNotFoundError(f"Файл не найден: {path}")

    try:
        with path.open("rb") as f:
            header = f.read(8)
    except OSError as exc:
        raise PresentationOpenError(f"Не удалось прочитать файл: {exc}") from exc

    if header.startswith(OLE2_MAGIC) or (
        path.suffix.lower() == ".ppt" and not zipfile.is_zipfile(path)
    ):
        raise PresentationOpenError(
            "Это старый формат PowerPoint 97–2003 (.ppt). "
            "Откройте файл в PowerPoint и пересохраните его: "
            "«Файл → Сохранить как → Презентация PowerPoint (*.pptx)»."
        )

    if not zipfile.is_zipfile(path):
        raise PresentationOpenError(
            "Файл повреждён или не является презентацией .pptx "
            "(архив OpenXML не распознан)."
        )

    corrupted_media: List[str] = []
    try:
        with zipfile.ZipFile(path, "r") as zf:
            infos = zf.infolist()
            bad_entries: List[str] = []
            for info in infos:
                try:
                    zf.read(info.filename)
                except Exception:
                    bad_entries.append(info.filename)

            if bad_entries:
                # Если повреждён XML слайда или структуры — файл не восстановить
                non_media = [
                    name for name in bad_entries
                    if not name.startswith("ppt/media/")
                ]
                if non_media:
                    raise PresentationOpenError(
                        f"Архив .pptx повреждён (битый внутренний файл «{non_media[0]}»). "
                        "Скачайте или сохраните презентацию заново."
                    )
                # Повреждены только медиафайлы — собираем копию в памяти для анализа слайдов
                corrupted_media = bad_entries
                mem_zip = BytesIO()
                with zipfile.ZipFile(mem_zip, "w") as zout:
                    for info in infos:
                        if info.filename in bad_entries:
                            zout.writestr(info.filename, b"")
                        else:
                            zout.writestr(info, zf.read(info.filename))
                mem_zip.seek(0)
                return Presentation(mem_zip), corrupted_media
    except PresentationOpenError:
        raise
    except zipfile.BadZipFile as exc:
        raise PresentationOpenError(
            "Архив .pptx повреждён. Скачайте презентацию заново."
        ) from exc

    try:
        return Presentation(path), []
    except (PackageNotFoundError, zipfile.BadZipFile, KeyError, ValueError, etree.XMLSyntaxError) as exc:
        raise PresentationOpenError(
            "Структура файла .pptx повреждена или это не презентация PowerPoint."
        ) from exc


# ------------------------------------------------------------------ проверки слайда

def check_corrupted_media(slide, slide_no: int, corrupted_media: Sequence[str],
                          reported_global: set, issues: List[Issue]) -> None:
    """1. Проверяет, ссылается ли слайд на повреждённый медиафайл внутри архива."""
    if not corrupted_media:
        return
    corrupted_norm = {"/" + m.lstrip("/") for m in corrupted_media}
    seen_on_slide = set()
    for rel in slide.part.rels.values():
        if getattr(rel, "is_external", False):
            continue
        try:
            partname = str(rel.target_part.partname)
        except Exception:
            continue
        if partname in corrupted_norm and partname not in seen_on_slide:
            seen_on_slide.add(partname)
            reported_global.add(partname)
            short_name = Path(partname).name
            issues.append(Issue(
                slide_no,
                "corrupted_media",
                f"Встроенный файл «{short_name}» повреждён в архиве — "
                f"на уроке вместо картинки будет пустая рамка",
                severity="error",
            ))


def check_flipped_images(nodes: Sequence[ShapeNode], slide_no: int,
                         issues: List[Issue]) -> None:
    """
    2. Картинка перевёрнута вверх ногами?
    Важно: горизонтальное отражение (flipH) — нормальный приём вёрстки (персонажи
    лицом друг к другу, симметричные стрелки/рамки), его не считаем ошибкой.
    Также учитываем, что flipV + поворот на 180° компенсируют друг друга.
    """
    for node in nodes:
        if _get_picture_blob(node.shape) is None and not (
            node.shape.shape_type == MSO_SHAPE_TYPE.PICTURE
        ):
            continue
        # Если включён flipV и картинка НЕ перевёрнута обратно поворотом на ~180°
        if node.flip_v and math.cos(math.radians(node.rotation_deg)) > 0.5:
            issues.append(Issue(
                slide_no,
                "flipped_image",
                f"{_shape_label(node.shape)} перевёрнута вверх ногами (отражение по вертикали)",
                severity="warning",
            ))


def check_bounds(nodes: Sequence[ShapeNode], slide_no: int,
                 slide_w: int, slide_h: int, issues: List[Issue]) -> None:
    """
    3. Выходит ли видимый объект за границы слайда?
    Отсеиваем микропогрешности (< 3 мм), пустые невидимые текстовые рамки
    и фоновые подложки слайда с небольшим выпуском под обрез.
    """
    max_bleed = int(1.5 * EMU_PER_CM)
    for node in nodes:
        shape = node.shape
        # Пустая текстовая рамка без заливки и границ невидима на экране
        if shape.shape_type in (MSO_SHAPE_TYPE.TEXT_BOX, MSO_SHAPE_TYPE.PLACEHOLDER):
            if _get_picture_blob(shape) is None and not _extract_shape_text(shape).strip():
                if not _has_visible_box_styling(shape):
                    continue

        left, top, right, bottom = _aabb(node)

        # Фоновая картинка на самом нижнем слое во весь слайд с небольшим вылетом — не баг
        if (
            node.z_order == (0,)
            and (shape.shape_type == MSO_SHAPE_TYPE.PICTURE or _get_picture_blob(shape) is not None)
            and (right - left) >= slide_w * 0.95
            and (bottom - top) >= slide_h * 0.95
            and left >= -max_bleed
            and top >= -max_bleed
            and right <= slide_w + max_bleed
            and bottom <= slide_h + max_bleed
        ):
            continue

        over = []
        if left < -OUT_OF_SLIDE_TOLERANCE:
            over.append(f"влево на {_cm(-left)}")
        if top < -OUT_OF_SLIDE_TOLERANCE:
            over.append(f"вверх на {_cm(-top)}")
        if right > slide_w + OUT_OF_SLIDE_TOLERANCE:
            over.append(f"вправо на {_cm(right - slide_w)}")
        if bottom > slide_h + OUT_OF_SLIDE_TOLERANCE:
            over.append(f"вниз на {_cm(bottom - slide_h)}")

        if over:
            issues.append(Issue(
                slide_no,
                "object_out_of_slide",
                f"{_shape_label(shape)} выходит за край слайда: " + ", ".join(over),
                severity="warning",
            ))


def _animated_shape_ids(slide) -> set:
    """ID объектов, участвующих в анимациях на слайде (интерактивные шторки/ответы)."""
    return {
        el.get("spid")
        for el in slide._element.findall(".//" + qn("p:spTgt"))
        if el.get("spid")
    }


def _is_region_opaque(pic_node: ShapeNode,
                      pic_rect: Tuple[float, float, float, float],
                      inter_rect: Tuple[float, float, float, float]) -> bool:
    """
    Проверяет, действительно ли картинка непрозрачна в области пересечения с текстом.
    Рамки и PNG-иллюстрации с прозрачным фоном не должны давать ложных тревог.
    """
    blob = _get_picture_blob(pic_node.shape)
    if not blob:
        return True
    try:
        with PILImage.open(BytesIO(blob)) as im:
            has_alpha = (
                "A" in im.getbands()
                or (im.mode == "P" and "transparency" in im.info)
            )
            if not has_alpha:
                return True
            rgba = im.convert("RGBA")
            w_px, h_px = rgba.size
            if w_px <= 0 or h_px <= 0:
                return True

            # Учитываем обрезку (crop) картинки в PowerPoint
            cl = float(getattr(pic_node.shape, "crop_left", 0.0) or 0.0)
            cr = float(getattr(pic_node.shape, "crop_right", 0.0) or 0.0)
            ct = float(getattr(pic_node.shape, "crop_top", 0.0) or 0.0)
            cb = float(getattr(pic_node.shape, "crop_bottom", 0.0) or 0.0)
            x0 = max(0, min(w_px - 1, int(w_px * cl)))
            y0 = max(0, min(h_px - 1, int(h_px * ct)))
            x1 = max(x0 + 1, min(w_px, int(w_px * (1.0 - cr))))
            y1 = max(y0 + 1, min(h_px, int(h_px * (1.0 - cb))))
            rgba = rgba.crop((x0, y0, x1, y1))

            if pic_node.flip_h:
                rgba = rgba.transpose(PILImage.Transpose.FLIP_LEFT_RIGHT)
            if pic_node.flip_v:
                rgba = rgba.transpose(PILImage.Transpose.FLIP_TOP_BOTTOM)
            rot_q = int(round(pic_node.rotation_deg / 90.0)) % 4
            if rot_q == 1:
                rgba = rgba.transpose(PILImage.Transpose.ROTATE_270)
            elif rot_q == 2:
                rgba = rgba.transpose(PILImage.Transpose.ROTATE_180)
            elif rot_q == 3:
                rgba = rgba.transpose(PILImage.Transpose.ROTATE_90)

            pw = pic_rect[2] - pic_rect[0]
            ph = pic_rect[3] - pic_rect[1]
            if pw <= 0 or ph <= 0:
                return True

            cw, ch = rgba.size
            rx0 = max(0, min(cw - 1, int(cw * (inter_rect[0] - pic_rect[0]) / pw)))
            ry0 = max(0, min(ch - 1, int(ch * (inter_rect[1] - pic_rect[1]) / ph)))
            rx1 = max(rx0 + 1, min(cw, int(cw * (inter_rect[2] - pic_rect[0]) / pw)))
            ry1 = max(ry0 + 1, min(ch, int(ch * (inter_rect[3] - pic_rect[1]) / ph)))

            sub = rgba.crop((rx0, ry0, rx1, ry1)).resize((24, 24))
            alpha_bytes = sub.getchannel("A").tobytes()
            opaque_ratio = sum(1 for a in alpha_bytes if a > 40) / len(alpha_bytes)
            return opaque_ratio >= 0.12
    except Exception:
        return True


def check_overlap(slide, nodes: Sequence[ShapeNode], slide_no: int,
                  issues: List[Issue]) -> None:
    """
    4. Перекрывает ли картинка текст СВЕРХУ?
    Проверяем:
      - порядок слоёв: картинка должна лежать ВЫШЕ текста (pic.z_order > txt.z_order),
        поэтому фоновые изображения и подложки не вызывают ложных срабатываний;
      - отсутствие интерактивной анимации появления/исчезновения;
      - заметный размер области пересечения (> 3 мм) и непрозрачность PNG в этом месте.
    """
    animated = _animated_shape_ids(slide)
    pictures = [n for n in nodes if _get_picture_blob(n.shape) is not None]
    text_nodes = [
        n for n in nodes
        if _extract_shape_text(n.shape).strip()
        and _get_picture_blob(n.shape) is None
    ]

    for pic in pictures:
        if str(pic.shape.shape_id) in animated:
            continue
        pr = _aabb(pic)
        for txt in text_nodes:
            if pic.z_order <= txt.z_order:
                continue
            if str(txt.shape.shape_id) in animated:
                continue
            tr = _aabb(txt)
            il = max(pr[0], tr[0])
            it = max(pr[1], tr[1])
            ir = min(pr[2], tr[2])
            ib = min(pr[3], tr[3])
            if (ir - il) <= OVERLAP_MIN_DIM or (ib - it) <= OVERLAP_MIN_DIM:
                continue
            if not _is_region_opaque(pic, pr, (il, it, ir, ib)):
                continue

            issues.append(Issue(
                slide_no,
                "overlap_picture_text",
                f"{_shape_label(pic.shape)} лежит поверх и закрывает {_shape_label(txt.shape)}",
                severity="error",
            ))
            break


def check_low_resolution(nodes: Sequence[ShapeNode], slide_no: int,
                         issues: List[Issue]) -> None:
    """
    5. Картинка сильно растянута и будет размытой на проекторе?
    Учитываем обрезку (crop) изображения и реальное экранное разрешение проекторов.
    """
    for node in nodes:
        if node.width <= 0 or node.height <= 0:
            continue
        blob = _get_picture_blob(node.shape)
        if not blob:
            continue
        try:
            with PILImage.open(BytesIO(blob)) as im:
                px_w, px_h = im.size
        except Exception:
            continue
        if px_w <= 0 or px_h <= 0:
            continue

        cl = float(getattr(node.shape, "crop_left", 0.0) or 0.0)
        cr = float(getattr(node.shape, "crop_right", 0.0) or 0.0)
        ct = float(getattr(node.shape, "crop_top", 0.0) or 0.0)
        cb = float(getattr(node.shape, "crop_bottom", 0.0) or 0.0)
        vis_w = max(0.01, 1.0 - cl - cr)
        vis_h = max(0.01, 1.0 - ct - cb)
        eff_px_w = px_w * vis_w
        eff_px_h = px_h * vis_h

        w_in = node.width / EMU_PER_INCH
        h_in = node.height / EMU_PER_INCH
        dpi = min(eff_px_w / w_in, eff_px_h / h_in)
        min_dpi = MIN_DPI_ICON if max(w_in, h_in) < 1.5 else MIN_DPI_MAIN

        if dpi < min_dpi:
            issues.append(Issue(
                slide_no,
                "low_resolution_image",
                f"{_shape_label(node.shape)}: низкое разрешение ~{dpi:.0f} DPI "
                f"({int(round(eff_px_w))}×{int(round(eff_px_h))} пикс. на "
                f"{_cm(node.width)}×{_cm(node.height)}) — на проекторе будут видны пиксели",
                severity="warning",
            ))


def _get_embedded_fonts(prs: Presentation) -> set:
    """Возвращает множество шрифтов (в нижнем регистре), встроенных внутрь файла .pptx."""
    embedded = set()
    for el in prs.part._element.findall(".//" + qn("p:embeddedFont")):
        font_el = el.find(qn("p:font"))
        if font_el is not None and font_el.get("typeface"):
            embedded.add(font_el.get("typeface").strip().lower())
    return embedded


def _get_slide_theme_fonts(slide) -> Tuple[Optional[str], Optional[str]]:
    """Определяет шрифты темы слайда: (major_font для заголовков, minor_font для текста)."""
    try:
        master = slide.slide_layout.slide_master
        for rel in master.part.rels.values():
            if rel.reltype.endswith("/theme"):
                theme_el = etree.fromstring(rel.target_part.blob)
                major_el = theme_el.find(".//" + qn("a:majorFont") + "/" + qn("a:latin"))
                minor_el = theme_el.find(".//" + qn("a:minorFont") + "/" + qn("a:latin"))
                major = major_el.get("typeface") if major_el is not None else None
                minor = minor_el.get("typeface") if minor_el is not None else None
                return major, minor
    except Exception:
        pass
    return None, None


def _resolve_typeface(raw_name: Optional[str], is_title: bool,
                      major_font: Optional[str], minor_font: Optional[str]) -> Optional[str]:
    """Превращает служебные ссылки темы (+mn-lt, +mj-lt) и пустые значения в реальное имя шрифта."""
    name = (raw_name or "").strip()
    if not name:
        name = (major_font if is_title else minor_font) or ""
    if name.startswith("+mj"):
        return major_font
    if name.startswith("+mn"):
        return minor_font
    if name.startswith("+"):
        return None
    return name or None


def _iter_shape_text_frames(shape):
    """Возвращает все текстовые фреймы объекта (включая ячейки таблицы)."""
    if getattr(shape, "has_text_frame", False):
        yield shape.text_frame
    if getattr(shape, "has_table", False):
        for row in shape.table.rows:
            for cell in row.cells:
                yield cell.text_frame


def check_fonts_and_sizes(slide, nodes: Sequence[ShapeNode], slide_no: int,
                          embedded_fonts: set, issues: List[Issue]) -> None:
    """
    6–7. Проверяет нестандартные невстроенные шрифты и слишком мелкий кегль (< 14 pt).
    Корректно расшифровывает шрифты темы (+mn-lt / +mj-lt), учитывает встроенные
    в презентацию шрифты и группирует предупреждения без дублей на одном слайде.
    """
    major_font, minor_font = _get_slide_theme_fonts(slide)
    bad_fonts_map: dict = {}    # font_name -> список объектов
    small_texts: List[Tuple[float, str]] = []

    for node in nodes:
        shape = node.shape
        is_title = False
        is_footer_ph = False
        if shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
            try:
                ph_type = shape.placeholder_format.type
                is_title = ph_type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)
                is_footer_ph = ph_type in FOOTER_PLACEHOLDERS
            except Exception:
                pass

        label = _shape_label(shape)
        for tf in _iter_shape_text_frames(shape):
            font_scale = 1.0
            body_pr = tf._txBody.find(qn("a:bodyPr")) if hasattr(tf, "_txBody") else None
            if body_pr is not None:
                autofit = body_pr.find(qn("a:normAutofit"))
                if autofit is not None and autofit.get("fontScale"):
                    try:
                        font_scale = int(autofit.get("fontScale")) / 100000.0
                    except ValueError:
                        font_scale = 1.0

            for para in tf.paragraphs:
                para_def_font = None
                p_pr = para._p.find(qn("a:pPr"))
                if p_pr is not None:
                    def_rpr = p_pr.find(qn("a:defRPr"))
                    if def_rpr is not None:
                        latin = def_rpr.find(qn("a:latin"))
                        if latin is not None:
                            para_def_font = latin.get("typeface")

                for run in para.runs:
                    if not run.text or not run.text.strip():
                        continue
                    raw_font = run.font.name or para_def_font
                    resolved = _resolve_typeface(raw_font, is_title, major_font, minor_font)
                    if resolved:
                        low = resolved.strip().lower()
                        if low not in STANDARD_FONTS and low not in embedded_fonts:
                            lst = bad_fonts_map.setdefault(resolved.strip(), [])
                            if label not in lst:
                                lst.append(label)

                    if not is_footer_ph and any(ch.isalnum() for ch in run.text):
                        if run.font.size is not None:
                            pt = round(run.font.size.pt * font_scale, 1)
                            if pt < MIN_FONT_PT:
                                small_texts.append((pt, label))
                        elif font_scale < 0.55:
                            # Текст без явного кегля сжат автоподгоном PowerPoint более чем вдвое
                            est_pt = round(24.0 * font_scale, 1)
                            small_texts.append((est_pt, label))

    for font_name in sorted(bad_fonts_map):
        targets = bad_fonts_map[font_name]
        where = targets[0] if len(targets) == 1 else f"{targets[0]} и ещё {len(targets) - 1} объ."
        issues.append(Issue(
            slide_no,
            "nonstandard_font",
            f"Шрифт «{font_name}» ({where}) не входит в стандартный набор Windows/Office "
            f"и не внедрён в файл — на школьном ПК может слететь вёрстка",
            severity="warning",
        ))

    if small_texts:
        min_pt, min_label = min(small_texts, key=lambda x: x[0])
        issues.append(Issue(
            slide_no,
            "small_font",
            f"{min_label}: очень мелкий текст ({min_pt:g} пт < {MIN_FONT_PT:g} пт) — "
            f"ученикам с задних парт будет плохо видно",
            severity="warning",
        ))


def check_auto_transition(slide, slide_no: int, issues: List[Issue]) -> None:
    """8. Забытый автопереход по таймеру или отключённый переход по клику."""
    tr = slide._element.find(qn("p:transition"))
    if tr is None:
        return
    if tr.get("advTm"):
        try:
            sec = int(tr.get("advTm")) / 1000.0
            sec_str = f"{sec:g} сек."
        except ValueError:
            sec_str = tr.get("advTm")
        issues.append(Issue(
            slide_no,
            "auto_advance_transition",
            f"Включён автопереход через {sec_str} — слайды будут уезжать сами во время урока",
            severity="error",
        ))
    elif tr.get("advClick") in ("0", "false"):
        issues.append(Issue(
            slide_no,
            "auto_advance_transition",
            "Отключено переключение слайда по клику мыши — на уроке слайд может «застрять»",
            severity="warning",
        ))


def _is_local_file_ref(ref: str, reltype: str) -> bool:
    low = ref.lower().strip()
    if low.startswith("file:"):
        return True
    if len(low) >= 3 and low[1] == ":" and low[2] in ("/", "\\"):
        return True
    if low.startswith(("\\\\", "../", "..\\", "./", ".\\")):
        return True
    if any(kind in reltype.lower() for kind in ("/video", "/audio", "/media", "/oleobject")):
        if not low.startswith(("http://", "https://")):
            return True
    return False


def check_external_links(slide, slide_no: int, issues: List[Issue]) -> None:
    """
    9. Внешние ссылки и невнедрённые локальные видео/аудио/документы.
    Почтовые ссылки mailto: игнорируются, одинаковые ссылки на одном слайде не дублируются.
    """
    rels = slide.part.rels
    items = list(rels.values()) if hasattr(rels, "values") else list(rels)
    seen = set()
    for rel in items:
        if not getattr(rel, "is_external", False):
            continue
        target = str(getattr(rel, "target_ref", "") or "").strip()
        if not target or target.lower().startswith("mailto:"):
            continue
        if target in seen:
            continue
        seen.add(target)

        reltype = str(getattr(rel, "reltype", "") or "")
        if _is_local_file_ref(target, reltype):
            issues.append(Issue(
                slide_no,
                "external_link",
                f"Ссылка на локальный файл «{target}» — файл не встроен в презентацию "
                f"и не откроется на чужом компьютере",
                severity="error",
            ))
        else:
            issues.append(Issue(
                slide_no,
                "external_link",
                f"Внешняя интернет-ссылка ({target}) — в классе без интернета не откроется",
                severity="warning",
            ))


# ------------------------------------------------------------------ автопочинка (--fix)

def _fix_shapes_recursive(shapes, slide_no: int, fixed: List[Issue],
                          parent_rot: float = 0.0,
                          parent_flip_v: bool = False) -> None:
    for shape in shapes:
        rot_deg, _, flip_v = _read_xfrm_attrs(shape)
        total_rot = (parent_rot + rot_deg) % 360.0
        eff_flip_v = parent_flip_v ^ flip_v

        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            _fix_shapes_recursive(shape.shapes, slide_no, fixed, total_rot, eff_flip_v)
            continue

        if _get_picture_blob(shape) is None and shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
            continue

        # Снимаем только flipV у перевёрнутых вверх ногами картинок (flipH трогать нельзя!)
        if eff_flip_v and math.cos(math.radians(total_rot)) > 0.5:
            sp_pr = getattr(shape._element, "spPr", None)
            xfrm = sp_pr.find(qn("a:xfrm")) if sp_pr is not None else None
            if xfrm is not None and xfrm.get("flipV") in ("1", "true"):
                del xfrm.attrib["flipV"]
                fixed.append(Issue(
                    slide_no,
                    "flipped_image",
                    f"{_shape_label(shape)}: убрано вертикальное переворачивание (flipV)",
                ))


def fix_presentation(src_path, dst_path=None) -> Tuple[List[Issue], Path]:
    """
    Исправляет объективно безопасные дефекты и сохраняет результат в КОПИЮ файла:
      - flipped_image: убирает переворот вверх ногами (flipV), НЕ ломая зеркальные flipH;
      - auto_advance_transition: отключает автотаймер advTm и включает переход по клику.
    Исходный файл никогда не изменяется.
    """
    src = Path(src_path)
    try:
        prs = Presentation(src)
    except Exception as exc:
        # Пробуем дать точное диагностическое сообщение
        _open_presentation_safely(src)
        raise PresentationOpenError(
            "Невозможно автоматически сохранить исправленную копию: архив .pptx повреждён."
        ) from exc

    fixed: List[Issue] = []
    for slide_no, slide in enumerate(prs.slides, 1):
        _fix_shapes_recursive(slide.shapes, slide_no, fixed)

        tr = slide._element.find(qn("p:transition"))
        if tr is not None:
            if tr.get("advTm"):
                del tr.attrib["advTm"]
                tr.set("advClick", "1")
                fixed.append(Issue(
                    slide_no,
                    "auto_advance_transition",
                    "Автопереход отключён — слайд переключается по клику",
                ))
            elif tr.get("advClick") in ("0", "false"):
                tr.set("advClick", "1")
                fixed.append(Issue(
                    slide_no,
                    "auto_advance_transition",
                    "Включено переключение слайда по клику мыши",
                ))

    if dst_path is None:
        dst_path = src.with_name(src.stem + "_исправлено" + src.suffix)
    dst = Path(dst_path)
    prs.save(dst)
    return fixed, dst


# ------------------------------------------------------------------ главные функции анализа

def inspect_presentation(path) -> InspectionResult:
    """Полная проверка презентации за одно чтение файла. Возвращает InspectionResult."""
    p = Path(path)
    prs, corrupted_media = _open_presentation_safely(p)
    embedded_fonts = _get_embedded_fonts(prs)

    issues: List[Issue] = []
    reported_corrupted: set = set()
    slide_w = int(prs.slide_width)
    slide_h = int(prs.slide_height)

    for slide_no, slide in enumerate(prs.slides, 1):
        nodes = list(_iter_slide_nodes(slide.shapes))
        check_corrupted_media(slide, slide_no, corrupted_media, reported_corrupted, issues)
        check_flipped_images(nodes, slide_no, issues)
        check_bounds(nodes, slide_no, slide_w, slide_h, issues)
        check_overlap(slide, nodes, slide_no, issues)
        check_low_resolution(nodes, slide_no, issues)
        check_fonts_and_sizes(slide, nodes, slide_no, embedded_fonts, issues)
        check_auto_transition(slide, slide_no, issues)
        check_external_links(slide, slide_no, issues)

    # Если битый медиафайл лежал в макете/теме, а не на конкретном слайде
    for media_path in corrupted_media:
        norm = "/" + media_path.lstrip("/")
        if norm not in reported_corrupted:
            issues.insert(0, Issue(
                1,
                "corrupted_media",
                f"Встроенный файл оформления «{Path(media_path).name}» повреждён в архиве .pptx",
                severity="error",
            ))

    return InspectionResult(
        path=p,
        slide_count=len(prs.slides),
        slide_width_cm=round(slide_w / EMU_PER_CM, 1),
        slide_height_cm=round(slide_h / EMU_PER_CM, 1),
        issues=issues,
    )


def check_presentation(path) -> List[Issue]:
    """Совместимая обёртка: возвращает список найденных проблем (List[Issue])."""
    return inspect_presentation(path).issues


# ------------------------------------------------------------------ форматирование отчёта

def format_report(path, issues: Sequence[Issue],
                  slide_count: Optional[int] = None,
                  slide_size_cm: Optional[Tuple[float, float]] = None) -> str:
    """Строит понятный текстовый отчёт для учителя (подходит и для консоли, и для окна)."""
    p = Path(path)
    if slide_count is None or slide_size_cm is None:
        try:
            prs, _ = _open_presentation_safely(p)
            slide_count = len(prs.slides)
            slide_size_cm = (
                round(int(prs.slide_width) / EMU_PER_CM, 1),
                round(int(prs.slide_height) / EMU_PER_CM, 1),
            )
        except Exception:
            slide_count = slide_count or 0

    lines: List[str] = []
    lines.append(f"ПРОВЕРКА ПРЕЗЕНТАЦИИ: {p.name}")
    size_info = (
        f" (формат {slide_size_cm[0]:g}×{slide_size_cm[1]:g} см)"
        if slide_size_cm else ""
    )
    lines.append(f"Путь: {p}")
    lines.append(f"Слайдов: {slide_count}{size_info}. Выполнено проверок: {len(IMPLEMENTED)}.")
    lines.append("-" * 68)

    if not issues:
        lines.append("[ОК] Проблем не найдено — презентация готова к показу на уроке!")
    else:
        by_slide: dict = {}
        for issue in issues:
            by_slide.setdefault(issue.slide, []).append(issue)

        for slide_no in sorted(by_slide):
            header = f"Слайд {slide_no}:" if slide_no > 0 else "Общие замечания к файлу:"
            lines.append(header)
            counts: dict = {}
            order: list = []
            for issue in by_slide[slide_no]:
                key = (issue.severity, issue.defect, issue.message)
                if key in counts:
                    counts[key] += 1
                else:
                    counts[key] = 1
                    order.append(key)
            for severity, defect, message in order:
                n = counts[(severity, defect, message)]
                tail = f" (×{n})" if n > 1 else ""
                badge = "[КРИТИЧНО]" if severity == "error" else "[ВНИМАНИЕ]"
                title = DEFECT_TITLES.get(defect, defect)
                lines.append(f"  {badge} {title}: {message}{tail}")

    lines.append("-" * 68)
    n_err = sum(1 for i in issues if i.severity == "error")
    n_warn = len(issues) - n_err
    if issues:
        lines.append(
            f"ИТОГО: проблем — {len(issues)} "
            f"(критичных: {n_err}, предупреждений: {n_warn})"
        )
        fixable = sum(1 for i in issues if i.defect in FIXABLE_DEFECTS)
        if fixable:
            lines.append(
                f"Подсказка: {fixable} из них можно исправить автоматически "
                f"(кнопка «Автоисправление» или флаг --fix)."
            )
        lines.append(
            "Рекомендация: поправьте указанные слайды или сохраните резервную "
            "копию в PDF на случай показа с чужого компьютера."
        )
    else:
        lines.append("ИТОГО: проблем — 0")
    return "\n".join(lines)


def print_report(path, issues: Sequence[Issue],
                 slide_count: Optional[int] = None,
                 slide_size_cm: Optional[Tuple[float, float]] = None) -> None:
    print(format_report(path, issues, slide_count=slide_count, slide_size_cm=slide_size_cm))


def format_batch_report(folder: Path) -> Tuple[str, int]:
    """Проверяет все презентации .pptx в папке и формирует сводную таблицу."""
    files = sorted(
        f for f in folder.glob("*.pptx")
        if not f.name.startswith("~$") and not f.name.endswith("_исправлено.pptx")
    )
    if not files:
        return f"В папке «{folder}» не найдено файлов .pptx.", 2

    lines = [
        f"СВОДНЫЙ ОТЧЁТ ПО ПАПКЕ: {folder}",
        f"Найдено презентаций: {len(files)}",
        "-" * 76,
        f"{'Файл':<40} | {'Слайдов':>7} | {'Критично':>8} | {'Внимание':>8} | {'Всего':>6}",
        "-" * 76,
    ]
    total_issues = 0
    for f in files:
        short = f.name if len(f.name) <= 40 else f.name[:37] + "..."
        try:
            res = inspect_presentation(f)
            total_issues += len(res.issues)
            lines.append(
                f"{short:<40} | {res.slide_count:>7} | "
                f"{res.error_count:>8} | {res.warning_count:>8} | {len(res.issues):>6}"
            )
        except PresentationOpenError as exc:
            total_issues += 1
            lines.append(f"{short:<40} | {'ОШИБКА':>7} | {str(exc)[:28]}")
    lines.append("-" * 76)
    lines.append(f"ИТОГО по папке ({len(files)} файлов): замечаний — {total_issues}")
    return "\n".join(lines), (1 if total_issues > 0 else 0)


# ------------------------------------------------------------------ самопроверка (--test)

def run_test() -> int:
    expected_path = TEST_FILES / "expected_defects.json"
    if not expected_path.exists():
        print("Не найден test_files/expected_defects.json. "
              "Сначала запустите: python3 generate_test_files.py")
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

        if not expected_list and not found:
            print("  ✅ чистый файл со сложной вёрсткой принят (ложных срабатываний нет)")

    print("\n" + "=" * 60)
    print("ИТОГ ТЕСТА:", "PASS ✅" if all_ok else "FAIL ❌")
    return 0 if all_ok else 1


# ------------------------------------------------------------------ CLI

def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description="Проверка презентации .pptx перед уроком"
    )
    parser.add_argument(
        "file",
        nargs="?",
        help="путь к презентации .pptx или к папке с презентациями",
    )
    parser.add_argument(
        "--test",
        action="store_true",
        help="самопроверка программы на эталонных тестовых файлах",
    )
    parser.add_argument(
        "--fix",
        action="store_true",
        help="безопасно исправить автопереходы и перевёрнутые картинки в копию файла",
    )
    args = parser.parse_args(argv)

    if args.test:
        return run_test()

    if not args.file:
        try:
            import app_ui
            return app_ui.run_app()
        except Exception:
            parser.print_help()
            print("\nГрафическое окно недоступно. Запустите из командной строки: см. выше.")
            return 2

    path = Path(args.file)
    if not path.exists():
        print(f"Файл или папка не найдены: {path}")
        return 2

    if path.is_dir():
        report_text, code = format_batch_report(path)
        print(report_text)
        return code

    if args.fix:
        try:
            before_res = inspect_presentation(path)
            fixed, dst = fix_presentation(path)
            after_res = inspect_presentation(dst)
        except PresentationOpenError as exc:
            print(f"❌ Не удалось обработать файл: {path}")
            print(f"   {exc}")
            return 3
        print(f"ИСХОДНИК НЕ ИЗМЕНЁН: {path}")
        print(f"Создана копия:       {dst}")
        print(f"Исправлено проблем:  {len(fixed)}")
        for f in fixed:
            print(f"  ✔ слайд {f.slide}: {f.message}")
        print(f"Проблем до: {len(before_res.issues)} → после: {len(after_res.issues)}")
        print()
        print_report(
            dst,
            after_res.issues,
            slide_count=after_res.slide_count,
            slide_size_cm=(after_res.slide_width_cm, after_res.slide_height_cm),
        )
        return 0

    try:
        res = inspect_presentation(path)
    except PresentationOpenError as exc:
        print(f"❌ Не удалось открыть файл: {path}")
        print(f"   {exc}")
        return 3

    print_report(
        path,
        res.issues,
        slide_count=res.slide_count,
        slide_size_cm=(res.slide_width_cm, res.slide_height_cm),
    )
    return 1 if res.issues else 0


if __name__ == "__main__":
    sys.exit(main())
