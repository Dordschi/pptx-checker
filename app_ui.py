#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
app_ui.py — графическое окно программы «Проверка презентации .pptx перед уроком».

Возможности:
  - выбор файла .pptx и мгновенная проверка в фоновом потоке (окно не зависает);
  - проверка целой папки с презентациями (сводная таблица);
  - кнопка «Автоисправление (в копию)» для безопасной починки перевёрнутых
    картинок и забытых автопереходов без изменения оригинала;
  - фильтрация замечаний («Все» / «Только критичные»);
  - цветовая подсветка уровней важности, прокрутка и защита отчёта от случайного стирания;
  - сохранение готового отчёта в текстовый файл (.txt).
"""

from __future__ import annotations

import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Optional

from check_presentation import (
    DEFECT_TITLES,
    FIXABLE_DEFECTS,
    IMPLEMENTED,
    InspectionResult,
    PresentationOpenError,
    fix_presentation,
    format_batch_report,
    format_report,
    inspect_presentation,
)


def _ui_font_family() -> str:
    if sys.platform.startswith("win"):
        return "Segoe UI"
    return "DejaVu Sans"


def _mono_font_family() -> str:
    if sys.platform.startswith("win"):
        return "Consolas"
    return "DejaVu Sans Mono"


class CheckerApp:
    """Главное окно приложения проверки презентаций."""

    def __init__(self, root: tk.Tk, initial_path: Optional[str] = None) -> None:
        self.root = root
        self.root.title("Проверка презентации .pptx перед уроком")
        self.root.geometry("880x620")
        self.root.minsize(720, 480)
        self.root.configure(bg="#f4f6f9")

        self.font_ui = _ui_font_family()
        self.font_mono = _mono_font_family()

        self.current_path: Optional[Path] = None
        self.last_result: Optional[InspectionResult] = None
        self.last_report_text: str = ""
        self.has_valid_report: bool = False
        self.only_critical_var = tk.BooleanVar(value=False)

        self._task_queue: queue.Queue = queue.Queue()
        self._busy: bool = False

        self._build_ui()
        self._set_welcome_text()

        if initial_path:
            p = Path(initial_path)
            if p.exists():
                self.root.after(120, lambda: self._load_path(p))
            else:
                self._set_banner(
                    f"Файл не найден: {p}",
                    bg="#fbe9eb",
                    fg="#9a1c28",
                )

    # ---------------------------------------------------------------- вёрстка окна

    def _build_ui(self) -> None:
        style = ttk.Style(self.root)
        try:
            if "clam" in style.theme_names():
                style.theme_use("clam")
        except Exception:
            pass

        style.configure("TButton", font=(self.font_ui, 10), padding=(10, 6))
        style.configure(
            "Primary.TButton",
            font=(self.font_ui, 10, "bold"),
            background="#1f6feb",
            foreground="#ffffff",
        )
        style.map(
            "Primary.TButton",
            background=[("active", "#1958b7"), ("disabled", "#a8b8d0")],
            foreground=[("disabled", "#eef2f7")],
        )
        style.configure(
            "Success.TButton",
            font=(self.font_ui, 10, "bold"),
            background="#2da44e",
            foreground="#ffffff",
        )
        style.map(
            "Success.TButton",
            background=[("active", "#23863e"), ("disabled", "#a9cbb2")],
            foreground=[("disabled", "#eef5f0")],
        )

        # Заголовок
        header = tk.Frame(self.root, bg="#ffffff", bd=0, highlightthickness=1,
                          highlightbackground="#dde3ec")
        header.pack(fill=tk.X, padx=14, pady=(12, 8))

        title_row = tk.Frame(header, bg="#ffffff")
        title_row.pack(fill=tk.X, padx=14, pady=(10, 4))

        tk.Label(
            title_row,
            text="Проверка презентации .pptx перед уроком",
            font=(self.font_ui, 14, "bold"),
            bg="#ffffff",
            fg="#1b2430",
        ).pack(side=tk.LEFT)

        self.file_label = tk.Label(
            header,
            text="Файл не выбран",
            anchor="w",
            font=(self.font_ui, 10),
            bg="#ffffff",
            fg="#57606a",
        )
        self.file_label.pack(fill=tk.X, padx=14, pady=(0, 10))

        # Панель кнопок
        toolbar = tk.Frame(self.root, bg="#f4f6f9")
        toolbar.pack(fill=tk.X, padx=14, pady=(0, 8))

        self.choose_btn = ttk.Button(
            toolbar, text="Выбрать .pptx…", style="Primary.TButton",
            command=self.choose_file,
        )
        self.choose_btn.pack(side=tk.LEFT)

        self.folder_btn = ttk.Button(
            toolbar, text="Проверить папку…",
            command=self.choose_folder,
        )
        self.folder_btn.pack(side=tk.LEFT, padx=(8, 0))

        self.rerun_btn = ttk.Button(
            toolbar, text="Проверить заново",
            command=self.rerun_current, state=tk.DISABLED,
        )
        self.rerun_btn.pack(side=tk.LEFT, padx=(8, 0))

        self.fix_btn = ttk.Button(
            toolbar, text="Автоисправление (в копию)…",
            style="Success.TButton",
            command=self.run_fix, state=tk.DISABLED,
        )
        self.fix_btn.pack(side=tk.LEFT, padx=(8, 0))

        self.save_btn = ttk.Button(
            toolbar, text="Сохранить отчёт…",
            command=self.save_report, state=tk.DISABLED,
        )
        self.save_btn.pack(side=tk.RIGHT)

        # Плашка статуса + фильтр
        status_row = tk.Frame(self.root, bg="#f4f6f9")
        status_row.pack(fill=tk.X, padx=14, pady=(0, 8))

        self.banner = tk.Label(
            status_row,
            text="Выберите файл .pptx, чтобы проверить презентацию перед показом в классе.",
            anchor="w",
            padx=12,
            pady=7,
            font=(self.font_ui, 10, "bold"),
            bg="#e8eef6",
            fg="#243b55",
        )
        self.banner.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.crit_check = tk.Checkbutton(
            status_row,
            text="Только критичные",
            variable=self.only_critical_var,
            command=self._refresh_filtered_view,
            font=(self.font_ui, 9),
            bg="#f4f6f9",
            activebackground="#f4f6f9",
        )
        self.crit_check.pack(side=tk.RIGHT, padx=(10, 0))

        # Поле отчёта с прокруткой
        report_frame = tk.Frame(
            self.root, bg="#ffffff", highlightthickness=1, highlightbackground="#cfd7e3"
        )
        report_frame.pack(fill=tk.BOTH, expand=True, padx=14, pady=(0, 14))

        scrollbar = ttk.Scrollbar(report_frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.output = tk.Text(
            report_frame,
            wrap=tk.WORD,
            font=(self.font_mono, 10),
            bg="#ffffff",
            fg="#1f2328",
            padx=14,
            pady=12,
            bd=0,
            yscrollcommand=scrollbar.set,
            state=tk.DISABLED,
        )
        self.output.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.output.yview)

        # Цветовые стили строк отчёта
        self.output.tag_configure("heading", font=(self.font_mono, 10, "bold"), foreground="#0f2942")
        self.output.tag_configure("slide", font=(self.font_mono, 10, "bold"), foreground="#1f6feb")
        self.output.tag_configure("error", foreground="#b42318", font=(self.font_mono, 10, "bold"))
        self.output.tag_configure("warning", foreground="#9a5b00")
        self.output.tag_configure("ok", foreground="#1a7f37", font=(self.font_mono, 10, "bold"))
        self.output.tag_configure("muted", foreground="#6e7781")

    # ---------------------------------------------------------------- управление текстом и статусом

    def _set_banner(self, text: str, bg: str = "#e8eef6", fg: str = "#243b55") -> None:
        self.banner.config(text=text, bg=bg, fg=fg)

    def _write_report_lines(self, text: str) -> None:
        self.output.config(state=tk.NORMAL)
        self.output.delete("1.0", tk.END)
        for line in text.splitlines():
            tag = None
            stripped = line.strip()
            if stripped.startswith(("ПРОВЕРКА ПРЕЗЕНТАЦИИ:", "СВОДНЫЙ ОТЧЁТ", "ИТОГО")):
                tag = "heading"
            elif stripped.startswith("Слайд ") or stripped.startswith("Общие замечания"):
                tag = "slide"
            elif "[КРИТИЧНО]" in line or line.startswith("❌"):
                tag = "error"
            elif "[ВНИМАНИЕ]" in line:
                tag = "warning"
            elif "[ОК]" in line or stripped.startswith("✔"):
                tag = "ok"
            elif stripped.startswith(("---", "Путь:", "Слайдов:", "Подсказка:", "Рекомендация:")):
                tag = "muted"

            if tag:
                self.output.insert(tk.END, line + "\n", tag)
            else:
                self.output.insert(tk.END, line + "\n")
        self.output.config(state=tk.DISABLED)

    def _set_welcome_text(self) -> None:
        welcome = (
            "Как пользоваться программой:\n"
            "  1. Нажмите «Выбрать .pptx…» и укажите учебную презентацию.\n"
            "  2. Программа проверит все слайды по 9 параметрам:\n"
            "     • перевёрнутые вверх ногами картинки;\n"
            "     • выход текста, фигур и картинок за края слайда;\n"
            "     • перекрытие текста непрозрачными картинками сверху;\n"
            "     • растянутые картинки низкого разрешения (размытие на проекторе);\n"
            "     • нестандартные невстроенные шрифты и слишком мелкий кегль (< 14 пт);\n"
            "     • забытые автопереходы по таймеру;\n"
            "     • внешние ссылки, невнедрённые локальные видео/аудио и битые медиафайлы.\n"
            "  3. При наличии перевёрнутых картинок или автопереходов нажмите\n"
            "     «Автоисправление (в копию)…» — исходный файл останется нетронутым.\n"
        )
        self._write_report_lines(welcome)

    def _set_busy(self, busy: bool, status_msg: Optional[str] = None) -> None:
        self._busy = busy
        if busy:
            self.choose_btn.config(state=tk.DISABLED)
            self.folder_btn.config(state=tk.DISABLED)
            self.rerun_btn.config(state=tk.DISABLED)
            self.fix_btn.config(state=tk.DISABLED)
            self.save_btn.config(state=tk.DISABLED)
            self.crit_check.config(state=tk.DISABLED)
            if status_msg:
                self._set_banner(status_msg, bg="#e8eef6", fg="#1f4e79")
        else:
            self.choose_btn.config(state=tk.NORMAL)
            self.folder_btn.config(state=tk.NORMAL)
            self.rerun_btn.config(state=tk.NORMAL if self.current_path else tk.DISABLED)
            self.save_btn.config(state=tk.NORMAL if self.has_valid_report else tk.DISABLED)
            self.crit_check.config(state=tk.NORMAL if self.last_result else tk.DISABLED)
            can_fix = bool(self.last_result and self.last_result.fixable_count > 0)
            self.fix_btn.config(state=tk.NORMAL if can_fix else tk.DISABLED)

    # ---------------------------------------------------------------- действия пользователя

    def choose_file(self) -> None:
        if self._busy:
            return
        path = filedialog.askopenfilename(
            title="Выберите презентацию PowerPoint",
            filetypes=[
                ("Презентация PowerPoint (*.pptx)", "*.pptx"),
                ("Старый формат PowerPoint (*.ppt)", "*.ppt"),
                ("Все файлы", "*.*"),
            ],
        )
        if path:
            self._load_path(Path(path))

    def choose_folder(self) -> None:
        if self._busy:
            return
        folder = filedialog.askdirectory(title="Выберите папку с презентациями .pptx")
        if folder:
            self._load_path(Path(folder))

    def rerun_current(self) -> None:
        if not self._busy and self.current_path:
            self._load_path(self.current_path)

    def _load_path(self, path: Path) -> None:
        self.current_path = path
        if path.is_dir():
            self.file_label.config(text=f"Папка: {path}", fg="#1f2328")
            self.last_result = None
            self._start_worker(
                lambda: ("batch", format_batch_report(path)),
                status_msg="Проверяем все презентации в выбранной папке…",
            )
        else:
            self.file_label.config(text=f"Файл: {path.name}  ({path.parent})", fg="#1f2328")
            self._start_worker(
                lambda: ("single", inspect_presentation(path)),
                status_msg=f"Проверяем «{path.name}»…",
            )

    def run_fix(self) -> None:
        if self._busy or not self.current_path or self.current_path.is_dir():
            return
        src = self.current_path
        default_dst = src.with_name(src.stem + "_исправлено" + src.suffix)
        dst_str = filedialog.asksaveasfilename(
            title="Куда сохранить исправленную копию презентации",
            defaultextension=".pptx",
            initialdir=str(src.parent),
            initialfile=default_dst.name,
            filetypes=[("Презентация PowerPoint (*.pptx)", "*.pptx")],
        )
        if not dst_str:
            return
        dst = Path(dst_str)
        if dst.resolve() == src.resolve():
            messagebox.showwarning(
                "Выберите другое имя",
                "Чтобы не повредить оригинал, сохраните исправленную презентацию в отдельный файл.",
            )
            return

        def _job():
            before_res = inspect_presentation(src)
            fixed, saved_path = fix_presentation(src, dst)
            after_res = inspect_presentation(saved_path)
            return ("fixed", (before_res, fixed, saved_path, after_res))

        self._start_worker(_job, status_msg="Создаём исправленную копию презентации…")

    def save_report(self) -> None:
        if not self.has_valid_report or not self.last_report_text.strip():
            messagebox.showinfo("Отчёт ещё не готов", "Сначала выберите и проверьте презентацию.")
            return
        stem = self.current_path.stem if self.current_path else "презентация"
        path = filedialog.asksaveasfilename(
            title="Сохранить отчёт проверки",
            defaultextension=".txt",
            initialfile=f"отчёт_{stem}.txt",
            filetypes=[("Текстовый файл (*.txt)", "*.txt")],
        )
        if path:
            Path(path).write_text(self.last_report_text, encoding="utf-8")
            messagebox.showinfo("Отчёт сохранён", f"Файл сохранён:\n{path}")

    def _refresh_filtered_view(self) -> None:
        if not self.last_result:
            return
        res = self.last_result
        issues = (
            [i for i in res.issues if i.severity == "error"]
            if self.only_critical_var.get()
            else res.issues
        )
        text = format_report(
            res.path,
            issues,
            slide_count=res.slide_count,
            slide_size_cm=(res.slide_width_cm, res.slide_height_cm),
        )
        if self.only_critical_var.get() and len(issues) < len(res.issues):
            hidden = len(res.issues) - len(issues)
            text += f"\n\n(Фильтр «Только критичные» активен: скрыто некритичных предупреждений — {hidden}.)"
        self.last_report_text = text
        self._write_report_lines(text)

    # ---------------------------------------------------------------- фоновый поток

    def _start_worker(self, fn, status_msg: str) -> None:
        self._set_busy(True, status_msg=status_msg)

        def _target():
            try:
                kind, payload = fn()
                self._task_queue.put(("ok", kind, payload))
            except Exception as exc:
                self._task_queue.put(("err", type(exc).__name__, exc))

        t = threading.Thread(target=_target, daemon=True)
        t.start()
        self.root.after(60, self._poll_worker)

    def _poll_worker(self) -> None:
        try:
            status, kind, payload = self._task_queue.get_nowait()
        except queue.Empty:
            self.root.after(60, self._poll_worker)
            return

        if status == "err":
            self.last_result = None
            self.has_valid_report = False
            self._set_busy(False)
            msg = str(payload)
            self._set_banner("Не удалось проверить файл", bg="#fbe9eb", fg="#9a1c28")
            self._write_report_lines(f"❌ Ошибка при открытии файла:\n{msg}")
            if isinstance(payload, PresentationOpenError):
                messagebox.showerror("Не удалось открыть презентацию", msg)
            else:
                messagebox.showerror("Ошибка проверки", f"Причина: {msg}")
            return

        if kind == "single":
            res: InspectionResult = payload
            self.last_result = res
            self.has_valid_report = True
            self._apply_single_result(res)
            self._set_busy(False)
        elif kind == "batch":
            report_text, _ = payload
            self.last_result = None
            self.has_valid_report = True
            self.last_report_text = report_text
            self._write_report_lines(report_text)
            self._set_banner("Сводная проверка папки завершена", bg="#e8eef6", fg="#1f4e79")
            self._set_busy(False)
        elif kind == "fixed":
            before_res, fixed, saved_path, after_res = payload
            self.current_path = saved_path
            self.file_label.config(
                text=f"Файл (исправленная копия): {saved_path.name}  ({saved_path.parent})",
                fg="#1a7f37",
            )
            self.last_result = after_res
            self.has_valid_report = True
            self.only_critical_var.set(False)

            header_lines = [
                f"✔ СОЗДАНА ИСПРАВЛЕННАЯ КОПИЯ: {saved_path.name}",
                f"Исходный файл не изменён: {before_res.path}",
                f"Автоматически устранено дефектов: {len(fixed)} "
                f"(было проблем: {len(before_res.issues)} → осталось: {len(after_res.issues)})",
            ]
            for item in fixed:
                title = DEFECT_TITLES.get(item.defect, item.defect)
                header_lines.append(f"  ✔ Слайд {item.slide} ({title}): {item.message}")
            header_lines.append("")

            body = format_report(
                after_res.path,
                after_res.issues,
                slide_count=after_res.slide_count,
                slide_size_cm=(after_res.slide_width_cm, after_res.slide_height_cm),
            )
            full_text = "\n".join(header_lines) + "\n" + body
            self.last_report_text = full_text
            self._write_report_lines(full_text)
            self._update_banner_for_result(after_res)
            self._set_busy(False)
            messagebox.showinfo(
                "Автоисправление выполнено",
                f"Исправлено проблем: {len(fixed)}.\n\n"
                f"Новая копия сохранена:\n{saved_path}",
            )

    def _update_banner_for_result(self, res: InspectionResult) -> None:
        if not res.issues:
            self._set_banner(
                f"Готово к уроку! Проблем не найдено (слайдов: {res.slide_count}, "
                f"формат {res.slide_width_cm:g}×{res.slide_height_cm:g} см).",
                bg="#dafbe1",
                fg="#116329",
            )
        elif res.error_count > 0:
            fix_hint = (
                f" • Автоисправимых: {res.fixable_count}"
                if res.fixable_count else ""
            )
            self._set_banner(
                f"Найдено замечаний: {len(res.issues)} "
                f"(критичных: {res.error_count}, предупреждений: {res.warning_count}){fix_hint}",
                bg="#fbe9eb",
                fg="#9a1c28",
            )
        else:
            fix_hint = (
                f" • Автоисправимых: {res.fixable_count}"
                if res.fixable_count else ""
            )
            self._set_banner(
                f"Найдено предупреждений: {len(res.issues)} (критичных ошибок нет){fix_hint}",
                bg="#fff4d6",
                fg="#7a4500",
            )

    def _apply_single_result(self, res: InspectionResult) -> None:
        self._update_banner_for_result(res)
        self._refresh_filtered_view()


def run_app(initial_path: Optional[str] = None) -> int:
    """Создаёт и запускает главное окно приложения."""
    root = tk.Tk()
    CheckerApp(root, initial_path=initial_path)
    root.mainloop()
    return 0


if __name__ == "__main__":
    arg_path = sys.argv[1] if len(sys.argv) > 1 else None
    sys.exit(run_app(initial_path=arg_path))
