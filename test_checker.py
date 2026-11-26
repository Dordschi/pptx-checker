#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модульные и регрессионные тесты ядра проверки презентаций (unittest).

Запуск:
  python3 -m unittest test_checker.py -v
"""

import tempfile
import unittest
import zipfile
from pathlib import Path

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from check_presentation import (
    OLE2_MAGIC,
    PresentationOpenError,
    check_presentation,
    fix_presentation,
    format_batch_report,
    inspect_presentation,
    run_test,
)

BASE = Path(__file__).resolve().parent
TEST_FILES = BASE / "test_files"


def _make_png(path: Path, size=(600, 400), color=(40, 90, 160, 255)) -> Path:
    img = Image.new("RGBA", size, color)
    img.save(path)
    return path


class TestPptxChecker(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.opaque_png = _make_png(self.tmp / "opaque.png", (800, 600), (30, 80, 150, 255))

        # Прозрачная рамка (в центре alpha = 0)
        frame = Image.new("RGBA", (800, 600), (0, 0, 0, 0))
        d = ImageDraw.Draw(frame)
        d.rectangle([0, 0, 799, 599], outline=(200, 50, 50, 255), width=10)
        self.transparent_frame = self.tmp / "frame.png"
        frame.save(self.transparent_frame)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def _new_prs(self) -> Presentation:
        prs = Presentation()
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(7.5)
        return prs

    def test_reference_suite_passes(self) -> None:
        """Эталонная самопроверка (--test) на broken.pptx и clean.pptx проходит на 100%."""
        self.assertEqual(run_test(), 0)

    def test_z_order_background_vs_foreground_overlap(self) -> None:
        """Фон под текстом (z=0 < z=1) — не ошибка; картинка поверх текста (z=1 > z=0) — ошибка."""
        prs = self._new_prs()
        blank = prs.slide_layouts[6]

        # Слайд 1: картинка СНИЗУ (фон), текст СВЕРХУ -> 0 проблем
        s1 = prs.slides.add_slide(blank)
        s1.shapes.add_picture(str(self.opaque_png), Inches(0), Inches(0), Inches(10), Inches(7.5))
        tb1 = s1.shapes.add_textbox(Inches(1), Inches(2), Inches(6), Inches(2))
        tb1.text_frame.text = "Заголовок поверх фона"
        tb1.text_frame.paragraphs[0].runs[0].font.size = Pt(28)

        # Слайд 2: текст СНИЗУ, непрозрачная картинка СВЕРХУ -> overlap_picture_text
        s2 = prs.slides.add_slide(blank)
        tb2 = s2.shapes.add_textbox(Inches(1), Inches(2), Inches(6), Inches(2))
        tb2.text_frame.text = "Скрытый картинкой текст"
        tb2.text_frame.paragraphs[0].runs[0].font.size = Pt(28)
        s2.shapes.add_picture(str(self.opaque_png), Inches(2), Inches(2), Inches(5), Inches(3))

        p = self.tmp / "zorder.pptx"
        prs.save(p)
        issues = check_presentation(p)
        self.assertEqual([(i.slide, i.defect) for i in issues], [(2, "overlap_picture_text")])
        self.assertEqual(issues[0].severity, "error")

    def test_transparent_png_frame_above_text_not_flagged(self) -> None:
        """PNG-рамка с прозрачной серединой поверх текста не считается перекрытием."""
        prs = self._new_prs()
        s = prs.slides.add_slide(prs.slide_layouts[6])
        tb = s.shapes.add_textbox(Inches(2), Inches(2.5), Inches(5), Inches(2))
        tb.text_frame.text = "Текст внутри прозрачной рамки"
        tb.text_frame.paragraphs[0].runs[0].font.size = Pt(24)
        s.shapes.add_picture(str(self.transparent_frame), Inches(1), Inches(1.5), Inches(7), Inches(4))

        p = self.tmp / "transparent.pptx"
        prs.save(p)
        self.assertEqual(check_presentation(p), [])

    def test_flip_h_preserved_and_flip_v_fixed(self) -> None:
        """Зеркальное отражение flipH не считается ошибкой и не ломается в --fix."""
        prs = self._new_prs()
        s = prs.slides.add_slide(prs.slide_layouts[6])
        pic_h = s.shapes.add_picture(str(self.opaque_png), Inches(0.5), Inches(1), Inches(3), Inches(2))
        pic_h._element.spPr.find(qn("a:xfrm")).set("flipH", "1")

        pic_v = s.shapes.add_picture(str(self.opaque_png), Inches(5.0), Inches(1), Inches(3), Inches(2))
        pic_v._element.spPr.find(qn("a:xfrm")).set("flipV", "1")

        # flipV + поворот на 180 градусов визуально стоит ровно -> не ошибка
        pic_v180 = s.shapes.add_picture(str(self.opaque_png), Inches(5.0), Inches(4), Inches(3), Inches(2))
        xfrm180 = pic_v180._element.spPr.find(qn("a:xfrm"))
        xfrm180.set("flipV", "1")
        xfrm180.set("rot", str(180 * 60000))

        src = self.tmp / "flips.pptx"
        prs.save(src)

        issues = check_presentation(src)
        self.assertEqual([(i.slide, i.defect) for i in issues], [(1, "flipped_image")])

        fixed, dst = fix_presentation(src)
        self.assertEqual(len(fixed), 1)
        self.assertEqual(check_presentation(dst), [])

        # Убеждаемся, что flipH=1 у первой картинки сохранился!
        prs_after = Presentation(dst)
        xfrm_h_after = prs_after.slides[0].shapes[0]._element.spPr.find(qn("a:xfrm"))
        self.assertEqual(xfrm_h_after.get("flipH"), "1")

    def test_theme_tokens_and_embedded_fonts(self) -> None:
        """Токены темы (+mn-lt) и встроенные в презентацию шрифты не дают ложных тревог."""
        prs = self._new_prs()
        s = prs.slides.add_slide(prs.slide_layouts[6])

        # 1) Токен темы +mn-lt
        tb1 = s.shapes.add_textbox(Inches(1), Inches(1), Inches(7), Inches(1))
        tb1.text_frame.text = "Текст со шрифтом темы"
        r1 = tb1.text_frame.paragraphs[0].runs[0]
        r1.font.name = "+mn-lt"
        r1.font.size = Pt(24)

        # 2) Редкий шрифт Lobster, НО он встроен в presentation.xml (<p:embeddedFontLst>)
        tb2 = s.shapes.add_textbox(Inches(1), Inches(2.5), Inches(7), Inches(1))
        tb2.text_frame.text = "Встроенный дизайнерский шрифт"
        r2 = tb2.text_frame.paragraphs[0].runs[0]
        r2.font.name = "Lobster"
        r2.font.size = Pt(24)

        emb_lst = prs.part._element.makeelement(qn("p:embeddedFontLst"), {})
        emb = prs.part._element.makeelement(qn("p:embeddedFont"), {})
        emb.append(prs.part._element.makeelement(qn("p:font"), {"typeface": "Lobster"}))
        emb_lst.append(emb)
        prs.part._element.append(emb_lst)

        # 3) Невстроенный редкий шрифт в нескольких полях -> ровно 1 сводное замечание на слайд
        for idx in range(3):
            tb = s.shapes.add_textbox(Inches(1 + idx * 2), Inches(4.5), Inches(1.5), Inches(1))
            tb.text_frame.text = f"Блок {idx}"
            run = tb.text_frame.paragraphs[0].runs[0]
            run.font.name = "Pacifico"
            run.font.size = Pt(22)

        p = self.tmp / "fonts.pptx"
        prs.save(p)
        issues = check_presentation(p)
        self.assertEqual([(i.slide, i.defect) for i in issues], [(1, "nonstandard_font")])
        self.assertIn("Pacifico", issues[0].message)

    def test_norm_autofit_shrunk_text_detected(self) -> None:
        """Текст 24 пт, сжатый автоподгоном PowerPoint (fontScale=45%) до 10.8 пт, ловится как small_font."""
        prs = self._new_prs()
        s = prs.slides.add_slide(prs.slide_layouts[6])
        tb = s.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(4))
        tb.text_frame.text = "Очень длинный текст, сжатый автоподгоном PowerPoint"
        tb.text_frame.paragraphs[0].runs[0].font.size = Pt(24)
        body_pr = tb.text_frame._txBody.find(qn("a:bodyPr"))
        body_pr.append(body_pr.makeelement(qn("a:normAutofit"), {"fontScale": "45000"}))

        p = self.tmp / "autofit.pptx"
        prs.save(p)
        issues = check_presentation(p)
        self.assertEqual([(i.slide, i.defect) for i in issues], [(1, "small_font")])

    def test_tables_and_groups_inspected(self) -> None:
        """Шрифты/кегль в таблицах и перевёрнутые картинки внутри групп обнаруживаются."""
        prs = self._new_prs()
        s = prs.slides.add_slide(prs.slide_layouts[6])

        # Таблица с мелким текстом 9 pt и нестандартным шрифтом
        tbl_shape = s.shapes.add_table(2, 2, Inches(0.5), Inches(0.5), Inches(4), Inches(2))
        cell = tbl_shape.table.cell(0, 0)
        cell.text = "Мелкий текст таблицы"
        run = cell.text_frame.paragraphs[0].runs[0]
        run.font.size = Pt(9)
        run.font.name = "Comfortaa"

        # Группа с перевёрнутой картинкой внутри
        grp = s.shapes.add_group_shape()
        pic = grp.shapes.add_picture(str(self.opaque_png), Inches(5.5), Inches(1), Inches(3), Inches(2))
        pic._element.spPr.find(qn("a:xfrm")).set("flipV", "1")
        # Инициализируем координаты группы
        grp.left, grp.top, grp.width, grp.height = Inches(5.5), Inches(1), Inches(3), Inches(2)
        grp_xfrm = grp._element.grpSpPr.find(qn("a:xfrm"))
        ch_off = grp_xfrm.find(qn("a:chOff"))
        ch_ext = grp_xfrm.find(qn("a:chExt"))
        ch_off.set("x", str(int(Inches(5.5))))
        ch_off.set("y", str(int(Inches(1))))
        ch_ext.set("cx", str(int(Inches(3))))
        ch_ext.set("cy", str(int(Inches(2))))

        p = self.tmp / "table_group.pptx"
        prs.save(p)
        defects = {i.defect for i in check_presentation(p)}
        self.assertEqual(defects, {"nonstandard_font", "small_font", "flipped_image"})

        fixed, dst = fix_presentation(p)
        self.assertEqual(len(fixed), 1)
        defects_after = {i.defect for i in check_presentation(dst)}
        self.assertNotIn("flipped_image", defects_after)

    def test_cropped_image_low_dpi_and_zero_height_safe(self) -> None:
        """Сильная обрезка (crop) большой картинки снижает эффективный DPI; высота 0 не роняет код."""
        prs = self._new_prs()
        s = prs.slides.add_slide(prs.slide_layouts[6])

        # 800x600 обрезаем на 90% (остаётся 80x60 пикс.) и растягиваем на 5x4 дюйма (~15 DPI)
        pic = s.shapes.add_picture(str(self.opaque_png), Inches(1), Inches(1), Inches(5), Inches(4))
        pic.crop_left = 0.45
        pic.crop_right = 0.45
        pic.crop_top = 0.45
        pic.crop_bottom = 0.45

        # Вырожденная картинка с нулевой высотой (защита от ZeroDivisionError)
        zero_pic = s.shapes.add_picture(str(self.opaque_png), Inches(7), Inches(1), Inches(2), Inches(1))
        zero_pic.height = 0

        p = self.tmp / "crop_zero.pptx"
        prs.save(p)
        issues = check_presentation(p)
        self.assertEqual([(i.slide, i.defect) for i in issues], [(1, "low_resolution_image")])

    def test_bounds_tolerance_and_rotation(self) -> None:
        """Микросмещение -24 EMU игнорируется, а повёрнутая на 90° длинная фигура ловится."""
        prs = self._new_prs()
        s = prs.slides.add_slide(prs.slide_layouts[6])

        tb_ok = s.shapes.add_textbox(Inches(1), Inches(0), Inches(5), Inches(1))
        tb_ok.top = -24
        tb_ok.text_frame.text = "Заголовок с погрешностью округления"
        tb_ok.text_frame.paragraphs[0].runs[0].font.size = Pt(24)

        # Пустая невидимая рамка за краем слайда игнорируется
        s.shapes.add_textbox(Inches(9), Inches(1), Inches(4), Inches(1))

        # Широкая плашка 6"x1" у верхнего края (center_y = 1.0"), при повороте на 90° верх уйдёт в -2.0"
        tb_rot = s.shapes.add_textbox(Inches(2), Inches(0.5), Inches(6), Inches(1))
        tb_rot.text_frame.text = "Повёрнутый блок"
        tb_rot.text_frame.paragraphs[0].runs[0].font.size = Pt(24)
        tb_rot.rotation = 90.0

        p = self.tmp / "bounds_rot.pptx"
        prs.save(p)
        issues = check_presentation(p)
        self.assertEqual([(i.slide, i.defect) for i in issues], [(1, "object_out_of_slide")])

    def test_external_links_classification(self) -> None:
        """mailto: игнорируется, локальный файл — error, веб-ссылка — warning."""
        prs = self._new_prs()
        s = prs.slides.add_slide(prs.slide_layouts[6])

        b_mail = s.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        b_mail.text_frame.text = "Почта"
        b_mail.text_frame.paragraphs[0].runs[0].font.size = Pt(20)
        b_mail.text_frame.paragraphs[0].runs[0].hyperlink.address = "mailto:info@school.ru"

        b_local = s.shapes.add_textbox(Inches(1), Inches(2.5), Inches(4), Inches(1))
        b_local.text_frame.text = "Локальное видео"
        b_local.text_frame.paragraphs[0].runs[0].font.size = Pt(20)
        b_local.text_frame.paragraphs[0].runs[0].hyperlink.address = "file:///C:/Users/Teacher/video.mp4"

        b_web = s.shapes.add_textbox(Inches(1), Inches(4.0), Inches(4), Inches(1))
        b_web.text_frame.text = "Сайт"
        b_web.text_frame.paragraphs[0].runs[0].font.size = Pt(20)
        b_web.text_frame.paragraphs[0].runs[0].hyperlink.address = "https://uchi.ru"

        p = self.tmp / "links.pptx"
        prs.save(p)
        issues = check_presentation(p)
        by_msg = {i.severity: i.message for i in issues}
        self.assertEqual(len(issues), 2)
        self.assertIn("локальный файл", by_msg["error"])
        self.assertIn("интернет-ссылка", by_msg["warning"])

    def test_ppt_binary_and_corrupted_archives(self) -> None:
        """Старый бинарный .ppt и битые файлы выдают понятные сообщения, а битая картинка в урок-1 локализуется."""
        fake_ppt = self.tmp / "old_lesson.ppt"
        fake_ppt.write_bytes(OLE2_MAGIC + b"\x00" * 128)
        with self.assertRaises(PresentationOpenError) as ctx:
            inspect_presentation(fake_ppt)
        self.assertIn("97–2003 (.ppt)", str(ctx.exception))

        garbage = self.tmp / "broken_zip.pptx"
        garbage.write_bytes(b"not a zip archive at all")
        with self.assertRaises(PresentationOpenError):
            inspect_presentation(garbage)

        # Реальная презентация с битым CRC у image1.png
        urok1 = TEST_FILES / "урок-1-повторение-нумерация-4-класс.pptx"
        if urok1.exists():
            res = inspect_presentation(urok1)
            self.assertEqual(res.slide_count, 5)
            self.assertTrue(
                any(i.defect == "corrupted_media" and i.slide == 1 for i in res.issues)
            )

    def test_real_presentations_have_no_legacy_false_positives(self) -> None:
        """На реальных учительских презентациях больше нет фейковых +mn-lt, -0.00\" и фоновых перекрытий."""
        report_text, _ = format_batch_report(TEST_FILES)
        self.assertIn("СВОДНЫЙ ОТЧЁТ ПО ПАПКЕ", report_text)

        for name in (
            "математика-3-класс-шаповалова.pptx",
            "повторение-3-класс-вспоминаем.pptx",
            "урок-11-диаграммы-4-класс.pptx",
            "урок-21-закрепление-4-класс.pptx",
            "урок-66-скорость-4-класс.pptx",
        ):
            p = TEST_FILES / name
            if not p.exists():
                continue
            issues = check_presentation(p)
            for issue in issues:
                self.assertNotIn("+mn-lt", issue.message)
                self.assertNotIn("-0,0 см", issue.message)
                self.assertNotEqual(issue.defect, "overlap_picture_text")
                self.assertNotEqual(issue.defect, "flipped_image")


if __name__ == "__main__":
    unittest.main()
