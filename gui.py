"""
GUI — главное окно приложения ЦСМС Отчёт
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from datetime import datetime

from app import (
    load_registry, save_registry,
    process_source, build_uid_data, build_excel,
)


# ─────────────────────────────────────────────
#  Окно управления справочником филиалов
# ─────────────────────────────────────────────
class RegistryWindow(tk.Toplevel):
    def __init__(self, parent, registry: dict, uids_without_filial: list):
        super().__init__(parent)
        self.title("Справочник: АТ → Филиал")
        self.geometry("560x480")
        self.resizable(False, True)
        self.registry = dict(registry)
        self.result   = None

        # ── Инструкция ──
        if uids_without_filial:
            msg = f"Новые АТ без филиала ({len(uids_without_filial)} шт.) — укажите филиал:"
            tk.Label(self, text=msg, fg="#C0392B", font=("Arial", 9, "bold"),
                     wraplength=530, justify="left").pack(padx=10, pady=(10, 4), anchor="w")
        else:
            tk.Label(self, text="Справочник привязки АТ к филиалам:",
                     font=("Arial", 9), justify="left").pack(padx=10, pady=(10, 4), anchor="w")

        # ── Таблица ──
        frame = tk.Frame(self)
        frame.pack(fill="both", expand=True, padx=10)

        cols = ("uid", "filial")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", height=18)
        self.tree.heading("uid",    text="Сетевой № АТ")
        self.tree.heading("filial", text="Филиал")
        self.tree.column("uid",    width=140, anchor="center")
        self.tree.column("filial", width=360, anchor="w")

        sb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # Подсветка новых
        self.tree.tag_configure("new", background="#FFF3CD")

        self._populate(uids_without_filial)

        # ── Поле редактирования ──
        edit_frame = tk.LabelFrame(self, text="Редактировать запись", padx=8, pady=6)
        edit_frame.pack(fill="x", padx=10, pady=6)

        tk.Label(edit_frame, text="АТ №:", width=8, anchor="e").grid(row=0, column=0)
        self.e_uid = tk.Entry(edit_frame, width=12)
        self.e_uid.grid(row=0, column=1, padx=4)

        tk.Label(edit_frame, text="Филиал:", width=8, anchor="e").grid(row=0, column=2)
        self.e_filial = tk.Entry(edit_frame, width=32)
        self.e_filial.grid(row=0, column=3, padx=4)

        tk.Button(edit_frame, text="Сохранить", command=self._save_entry,
                  bg="#27AE60", fg="white", padx=10).grid(row=0, column=4, padx=6)

        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        # ── Кнопки ──
        btn_frame = tk.Frame(self)
        btn_frame.pack(fill="x", padx=10, pady=(0, 10))
        tk.Button(btn_frame, text="✓  Применить и закрыть",
                  command=self._apply, bg="#2980B9", fg="white",
                  font=("Arial", 9, "bold"), padx=12, pady=4).pack(side="right")
        tk.Button(btn_frame, text="Отмена", command=self.destroy,
                  padx=10, pady=4).pack(side="right", padx=6)

    def _populate(self, highlight_uids):
        self.tree.delete(*self.tree.get_children())
        # Сначала новые
        new_set = {str(u) for u in highlight_uids}
        for uid, filial in sorted(self.registry.items(), key=lambda x: (str(x[0]) not in new_set, x[0])):
            tag = "new" if str(uid) in new_set else ""
            self.tree.insert("", "end", values=(uid, filial or ""), tags=(tag,))
        # Добавляем новые без записи
        for uid in highlight_uids:
            if str(uid) not in self.registry:
                self.tree.insert("", "end", values=(uid, ""), tags=("new",))

    def _on_select(self, _event):
        sel = self.tree.selection()
        if sel:
            uid, filial = self.tree.item(sel[0], "values")
            self.e_uid.delete(0, "end"); self.e_uid.insert(0, uid)
            self.e_filial.delete(0, "end"); self.e_filial.insert(0, filial)

    def _save_entry(self):
        uid    = self.e_uid.get().strip()
        filial = self.e_filial.get().strip()
        if not uid:
            messagebox.showwarning("Ошибка", "Введите номер АТ", parent=self)
            return
        if not filial:
            messagebox.showwarning("Ошибка", "Введите название филиала", parent=self)
            return
        self.registry[uid] = filial
        self._populate([])
        messagebox.showinfo("Сохранено", f"АТ {uid} → {filial}", parent=self)

    def _apply(self):
        self.result = self.registry
        self.destroy()


# ─────────────────────────────────────────────
#  Главное окно
# ─────────────────────────────────────────────
class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ЦСМС — Генератор отчёта")
        self.geometry("540x380")
        self.resizable(False, False)
        self.configure(bg="#F0F4F8")

        self.registry   = load_registry()
        self.source_path = tk.StringVar()
        self.status_var  = tk.StringVar(value="Готов к работе")

        self._build_ui()

    def _build_ui(self):
        # ── Заголовок ──
        hdr = tk.Frame(self, bg="#1F3864", height=56)
        hdr.pack(fill="x")
        tk.Label(hdr, text="ЦСМС  —  Генератор отчёта",
                 bg="#1F3864", fg="white",
                 font=("Arial", 14, "bold")).pack(side="left", padx=20, pady=10)

        # ── Шаг 1: выбор файла ──
        step1 = tk.LabelFrame(self, text=" Шаг 1 — Выберите исходный файл выгрузки ",
                               bg="#F0F4F8", font=("Arial", 9, "bold"), padx=10, pady=8)
        step1.pack(fill="x", padx=16, pady=(14, 4))

        tk.Entry(step1, textvariable=self.source_path, state="readonly",
                 width=50, relief="sunken").pack(side="left", padx=(0, 8))
        tk.Button(step1, text="Обзор...", command=self._browse,
                  padx=10).pack(side="left")

        # ── Шаг 2: справочник ──
        step2 = tk.LabelFrame(self, text=" Шаг 2 — Справочник филиалов ",
                               bg="#F0F4F8", font=("Arial", 9, "bold"), padx=10, pady=8)
        step2.pack(fill="x", padx=16, pady=4)

        self.lbl_registry = tk.Label(step2,
            text=f"Загружен: {len(self.registry)} АТ в справочнике",
            bg="#F0F4F8", font=("Arial", 9))
        self.lbl_registry.pack(side="left")
        tk.Button(step2, text="Открыть справочник",
                  command=self._open_registry, padx=10).pack(side="right")

        # ── Шаг 3: генерация ──
        step3 = tk.LabelFrame(self, text=" Шаг 3 — Сформировать отчёт ",
                               bg="#F0F4F8", font=("Arial", 9, "bold"), padx=10, pady=12)
        step3.pack(fill="x", padx=16, pady=4)

        self.btn_run = tk.Button(step3, text="▶  Сформировать СОГЛАСОВАН",
                                 command=self._run,
                                 bg="#1F6AA5", fg="white",
                                 font=("Arial", 10, "bold"),
                                 padx=20, pady=6, relief="flat", cursor="hand2")
        self.btn_run.pack()

        # ── Прогресс ──
        self.progress = ttk.Progressbar(self, mode="indeterminate", length=508)
        self.progress.pack(padx=16, pady=(8, 4))

        # ── Статус ──
        tk.Label(self, textvariable=self.status_var,
                 bg="#F0F4F8", fg="#555555",
                 font=("Arial", 9), wraplength=510).pack(padx=16)

        # ── Лог ──
        log_frame = tk.Frame(self, bg="#F0F4F8")
        log_frame.pack(fill="both", expand=True, padx=16, pady=(6, 12))
        self.log = tk.Text(log_frame, height=6, state="disabled",
                           font=("Courier", 8), bg="#1E1E1E", fg="#CCCCCC",
                           relief="flat", wrap="word")
        sb = ttk.Scrollbar(log_frame, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        self.log.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    def _log(self, msg: str, color: str = "#CCCCCC"):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")
        self.status_var.set(msg)

    def _browse(self):
        path = filedialog.askopenfilename(
            title="Выберите исходный файл выгрузки",
            filetypes=[("Excel файлы", "*.xlsx *.xls"), ("Все файлы", "*.*")]
        )
        if path:
            self.source_path.set(path)
            self._log(f"Файл выбран: {Path(path).name}")

    def _open_registry(self, uids_without_filial=None):
        win = RegistryWindow(self, self.registry, uids_without_filial or [])
        self.wait_window(win)
        if win.result is not None:
            self.registry = win.result
            save_registry(self.registry)
            self.lbl_registry.config(
                text=f"Загружен: {len(self.registry)} АТ в справочнике")
            self._log(f"Справочник сохранён ({len(self.registry)} АТ)")
        return win.result

    def _run(self):
        if not self.source_path.get():
            messagebox.showwarning("Нет файла", "Сначала выберите исходный файл выгрузки.")
            return
        self.btn_run.config(state="disabled")
        self.progress.start(12)
        threading.Thread(target=self._run_worker, daemon=True).start()

    def _run_worker(self):
        try:
            src = self.source_path.get()
            self._log("Читаю исходный файл...")
            df, month, year = process_source(src)
            self._log(f"Период: {month:02d}.{year} | ГЗ-терминалов: {df['UID'].nunique()}")

            self._log("Вычисляю активные дни с учётом ГЗ/ПДД...")
            uid_data = build_uid_data(df, month, year)

            # Проверяем новые АТ без филиала
            uids_without = [uid for uid in uid_data if str(uid) not in self.registry]
            if uids_without:
                self._log(f"⚠ Новых АТ без филиала: {len(uids_without)} — открываю справочник...")
                self.after(0, lambda: self._open_registry(uids_without))
                # Ждём пока пользователь закроет справочник
                import time
                time.sleep(1.5)

            # Куда сохранить
            src_stem = Path(src).stem
            default_name = f"ЦСМС_{year}_{month:02d}_СОГЛАСОВАН.xlsx"
            out_path = filedialog.asksaveasfilename(
                title="Сохранить отчёт как...",
                initialfile=default_name,
                defaultextension=".xlsx",
                filetypes=[("Excel файлы", "*.xlsx")]
            )
            if not out_path:
                self._log("Отменено пользователем.")
                return

            self._log("Формирую Excel...")
            build_excel(uid_data, self.registry, month, year, out_path)

            self._log(f"✓ Готово! Файл сохранён: {Path(out_path).name}")
            self._log(f"  АТ в отчёте: {len(uid_data)} | Период: {month:02d}.{year}")
            messagebox.showinfo(
                "Готово",
                f"Отчёт сформирован успешно!\n\n"
                f"АТ в отчёте: {len(uid_data)}\n"
                f"Период: {month:02d}.{year}\n\n"
                f"Файл: {Path(out_path).name}"
            )
            # Открыть папку
            os.startfile(str(Path(out_path).parent))

        except Exception as e:
            self._log(f"✗ Ошибка: {e}")
            messagebox.showerror("Ошибка", str(e))
        finally:
            self.progress.stop()
            self.btn_run.config(state="normal")


# ─────────────────────────────────────────────
#  Точка входа
# ─────────────────────────────────────────────
def main():
    app = MainWindow()
    app.mainloop()


if __name__ == "__main__":
    main()
