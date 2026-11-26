#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Точка входа для готовых сборок (exe / бинарник из GitHub Releases).

Двойной клик        → сразу открывается окно программы (без терминала).
С файлом-аргументом → окно открывается и сразу проверяет презентацию или папку:
    pptx-checker-windows.exe презентация.pptx

Сборка (выполняется в GitHub Actions):
    pyinstaller --onefile --windowed --name pptx-checker launcher.py
"""

import sys

import app_ui


def main() -> int:
    initial = sys.argv[1] if len(sys.argv) > 1 else None
    return app_ui.run_app(initial_path=initial)


if __name__ == "__main__":
    sys.exit(main())
