#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Первый набросок интерфейса «Проверка презентации перед уроком» (неделя 3).

Запуск:  python3 app_ui.py
Важно:  рядом должен лежать check_presentation.py (наше ядро с проверками).

Что умеет (пока набросок):
  - выбрать файл .pptx кнопкой «Обзор»
  - запустить проверку и показать отчёт в окне
  - сохранить отчёт в текстовый файл
  - обработка ошибок (не тот формат, повреждённый файл и т.д.)
"""

import io
import tkinter as tk
from contextlib import redirect_stdout
from pathlib import Path
from tkinter import filedialog, messagebox

from check_presentation import check_presentation, print_report

current_file = None  # путь к выбранной презентации


def choose_file():
    """Диалог выбора файла .pptx."""
    global current_file
    path = filedialog.askopenfilename(
        title="Выберите презентацию",
        filetypes=[("Презентация PowerPoint", "*.pptx"), ("Все файлы", "*.*")],
    )
    if path:
        current_file = path
        file_label.config(text=Path(path).name, fg="#1a7f37")
        run_btn.config(state=tk.NORMAL)


def run_check():
    """Запуск проверки и вывод отчёта в окно."""
    if not current_file:
        messagebox.showwarning("Файл не выбран",
                              "Сначала нажмите кнопку «Обзор»")
        return
    try:
        issues = check_presentation(current_file)
        buf = io.StringIO()
        with redirect_stdout(buf):          # перехватываем печать в строку
            print_report(current_file, issues)
        output.delete("1.0", tk.END)
        output.insert(tk.END, buf.read())
    except Exception as exc:                # ловим любые ошибки файла
        messagebox.showerror(
            "Не получилось проверить файл",
            f"Причина: {exc}\n\n"
            "Если файл старого формата .ppt — сохраните его как .pptx.")


def save_report():
    """Сохранить отчёт в текстовый файл."""
    text = output.get("1.0", tk.END).strip()
    if not text:
        messagebox.showinfo("Пусто", "Сначала выполните проверку")
        return
    path = filedialog.asksaveasfilename(
        title="Сохранить отчёт",
        defaultextension=".txt",
        initialfile="отчёт_проверки.txt",
        filetypes=[("Текстовый файл", "*.txt")],
    )
    if path:
        Path(path).write_text(text, encoding="utf-8")
        messagebox.showinfo("Готово",
                            f"Отчёт сохранён: {Path(path).name}")


# ---------------------------------------------------------------- окно
root = tk.Tk()
root.title("Проверка презентации перед уроком")
root.geometry("760x520")

tk.Label(root, text="Проверка презентации .pptx",
         font=("Arial", 14, "bold")).pack(pady=(12, 4))

frame = tk.Frame(root)
frame.pack(fill=tk.X, padx=16)

file_label = tk.Label(frame, text="файл не выбран",
                      anchor="w", font=("Arial", 11))
file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

choose_btn = tk.Button(frame, text="Обзор…", command=choose_file)
choose_btn.pack(side=tk.RIGHT, padx=(8, 0))

btns = tk.Frame(root)
btns.pack(fill=tk.X, padx=16, pady=10)

run_btn = tk.Button(btns, text="Проверить", command=run_check,
                    state=tk.DISABLED, bg="#2da44e", fg="white",
                    font=("Arial", 11, "bold"))
run_btn.pack(side=tk.LEFT)

save_btn = tk.Button(btns, text="Сохранить отчёт", command=save_report)
save_btn.pack(side=tk.LEFT, padx=10)

output = tk.Text(root, wrap=tk.WORD, font=("Consolas", 10),
                 state=tk.NORMAL, bg="#0d1117", fg="#c9d1d9")
output.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 12))
output.insert(tk.END, "Здесь появится отчёт: номера слайдов и найденные проблемы.\n")

if __name__ == "__main__":
    root.mainloop()
