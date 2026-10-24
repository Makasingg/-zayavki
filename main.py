import json
import os
import sys
import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

TYPES = ["Плановая", "Неплановая", "Аварийная", "Технологическое нарушение"]
STATUSES = ["Открыта", "В работе", "Закрыта"]
FIELDS = [
    ("date", "Дата"), ("type", "Тип работ"), ("object", "Объект"),
    ("equipment", "Оборудование"), ("description", "Описание"),
    ("responsible", "Ответственный"), ("start", "Начало"),
    ("end", "Окончание"), ("status", "Статус"),
]
WIDTHS = {"date": 90, "type": 170, "object": 130, "equipment": 130, "description": 260,
          "responsible": 130, "start": 60, "end": 70, "status": 90}


def data_path():
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "zayavki.json")


def today():
    return datetime.date.today().strftime("%d.%m.%Y")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Заявки на работы")
        self.geometry("1150x680")
        self.records = self.load()
        self.build_ui()
        self.refresh()

    def load(self):
        try:
            with open(data_path(), encoding="utf-8") as f:
                return json.load(f)
        except (OSError, ValueError):
            return []

    def save(self):
        with open(data_path(), "w", encoding="utf-8") as f:
            json.dump(self.records, f, ensure_ascii=False, indent=2)

    def build_ui(self):
        form = ttk.LabelFrame(self, text="Новая заявка", padding=10)
        form.pack(fill="x", padx=10, pady=8)
        self.vars = {}
        for i, (key, label) in enumerate(FIELDS):
            r, c = divmod(i, 3)
            ttk.Label(form, text=label).grid(row=r * 2, column=c, sticky="w", padx=6)
            if key == "type":
                w = ttk.Combobox(form, values=TYPES, state="readonly")
            elif key == "status":
                w = ttk.Combobox(form, values=STATUSES, state="readonly")
            else:
                w = ttk.Entry(form)
            w.grid(row=r * 2 + 1, column=c, sticky="ew", padx=6, pady=(0, 6))
            self.vars[key] = w
        for c in range(3):
            form.columnconfigure(c, weight=1)
        self.vars["type"].set(TYPES[0])
        self.vars["status"].set(STATUSES[0])
        self.clear()

        btns = ttk.Frame(form)
        btns.grid(row=6, column=0, columnspan=3, sticky="e", pady=6)
        ttk.Button(btns, text="Добавить", command=self.add).pack(side="left", padx=4)
        ttk.Button(btns, text="Очистить поля", command=self.clear).pack(side="left", padx=4)

        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=10)
        ttk.Button(bar, text="Сменить статус выбранной", command=self.change_status).pack(side="left")
        ttk.Button(bar, text="Удалить выбранную", command=self.delete).pack(side="left", padx=6)
        ttk.Button(bar, text="Выгрузить в Excel", command=self.export).pack(side="right")

        cols = [k for k, _ in FIELDS]
        self.tree = ttk.Treeview(self, columns=cols, show="headings", selectmode="browse")
        for k, label in FIELDS:
            self.tree.heading(k, text=label)
            self.tree.column(k, width=WIDTHS[k], anchor="w")
        self.tree.pack(fill="both", expand=True, padx=10, pady=8)

    def clear(self):
        for w in self.vars.values():
            if isinstance(w, ttk.Entry):
                w.delete(0, "end")
        self.vars["date"].insert(0, today())

    def add(self):
        rec = {k: self.vars[k].get().strip() for k, _ in FIELDS}
        if not rec["object"] or not rec["description"]:
            messagebox.showwarning("Нет данных", "Заполните поля «Объект» и «Описание».")
            return
        self.records.append(rec)
        self.save()
        self.refresh()
        self.clear()

    def selected_index(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Выбор", "Сначала выберите заявку в таблице.")
            return None
        return self.tree.index(sel[0])

    def change_status(self):
        i = self.selected_index()
        if i is None:
            return
        cur = self.records[i].get("status", "")
        nxt = STATUSES[(STATUSES.index(cur) + 1) % len(STATUSES)] if cur in STATUSES else STATUSES[0]
        self.records[i]["status"] = nxt
        self.save()
        self.refresh()

    def delete(self):
        i = self.selected_index()
        if i is None:
            return
        if messagebox.askyesno("Удаление", "Удалить выбранную заявку?"):
            del self.records[i]
            self.save()
            self.refresh()

    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        for r in self.records:
            self.tree.insert("", "end", values=[r.get(k, "") for k, _ in FIELDS])

    def export(self):
        if not self.records:
            messagebox.showinfo("Пусто", "Нет заявок для выгрузки.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
            initialfile=f"zayavki_{datetime.date.today():%Y-%m-%d}.xlsx")
        if not path:
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Заявки"
        headers = ["№"] + [label for _, label in FIELDS]
        ws.append(headers)
        for n, r in enumerate(self.records, 1):
            ws.append([n] + [r.get(k, "") for k, _ in FIELDS])

        thin = Side(style="thin", color="999999")
        border = Border(top=thin, bottom=thin, left=thin, right=thin)
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            for cell in row:
                cell.border = border
                if cell.row > 1:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
        for col in range(1, len(headers) + 1):
            ws.column_dimensions[get_column_letter(col)].width = 18
        ws.column_dimensions["A"].width = 6
        ws.column_dimensions[get_column_letter(headers.index("Описание") + 1)].width = 50
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        summary = wb.create_sheet("Сводка")
        summary.append(["Тип работ", "Количество"])
        for cell in summary[1]:
            cell.font = Font(bold=True)
        for t in TYPES:
            summary.append([t, sum(1 for r in self.records if r.get("type") == t)])
        summary.append(["Всего", len(self.records)])
        summary.column_dimensions["A"].width = 30
        summary.column_dimensions["B"].width = 14

        try:
            wb.save(path)
        except OSError as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")
            return
        messagebox.showinfo("Готово", f"Выгружено заявок: {len(self.records)}\n{path}")


if __name__ == "__main__":
    App().mainloop()
