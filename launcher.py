#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Точка входа для готовых сборок (exe / бинарник из GitHub Releases).

Двойной клик        → сразу открывается окно программы (без терминала).
С файлом-аргументом → окно открывается и сразу проверяет файл:
    pptx-checker-windows.exe презентация.pptx

Сборка (делает GitHub Actions):
    pyinstaller --onefile --windowed --name pptx-checker launcher.py
"""

import sys
from pathlib import Path

import app_ui


def preload_file_if_given() -> None:
    """Если передан путь к .pptx — подставить его в окно и сразу проверить."""
    if len(sys.argv) < 2:
        return
    path = sys.argv[1]
    if not Path(path).exists():
        return
    app_ui.current_file = path
    app_ui.file_label.config(text=Path(path).name, fg="#1a7f37")
    app_ui.run_btn.config(state="normal")
    app_ui.root.after(150, app_ui.run_check)


if __name__ == "__main__":
    preload_file_if_given()
    app_ui.root.mainloop()
