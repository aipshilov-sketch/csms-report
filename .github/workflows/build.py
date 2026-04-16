"""
Сборка .exe через PyInstaller.
Запускать на Windows: python build.py
"""
import subprocess
import sys

cmd = [
    sys.executable, "-m", "PyInstaller",
    "--onefile",                    # один .exe файл
    "--windowed",                   # без консоли
    "--name", "ЦСМС_Отчёт",
    "--icon", "icon.ico",           # иконка (если есть)
    "--add-data", "icon.ico;.",     # добавить иконку в пакет
    "--hidden-import", "openpyxl",
    "--hidden-import", "pandas",
    "--hidden-import", "tkinter",
    "gui.py",
]

# Убираем иконку если файла нет
import os
if not os.path.exists("icon.ico"):
    cmd = [c for c in cmd if "icon" not in c.lower()]

result = subprocess.run(cmd, capture_output=False)
if result.returncode == 0:
    print("\n✓ Сборка успешна! Файл: dist/ЦСМС_Отчёт.exe")
else:
    print("\n✗ Ошибка сборки")
    sys.exit(1)
