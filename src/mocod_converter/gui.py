import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, ttk
from zipfile import BadZipFile

from openpyxl.utils.exceptions import InvalidFileException

from .converter import ConversionError, convert


def open_folder(path):
    if sys.platform == "win32":
        os.startfile(path)
        return
    command = "open" if sys.platform == "darwin" else "xdg-open"
    subprocess.Popen([command, str(path)])


def main():
    root = tk.Tk()
    root.title("Реестр займов МФО")
    root.minsize(610, 305)
    root.configure(bg="#edf2f5")
    style = ttk.Style(root)
    style.configure("Card.TFrame", background="#ffffff")
    style.configure("Card.TLabel", background="#ffffff", foreground="#163047", font=("Segoe UI", 10))
    style.configure("Title.TLabel", background="#ffffff", foreground="#163047", font=("Segoe UI", 19, "bold"))
    style.configure("Action.TButton", padding=(18, 8), font=("Segoe UI", 10))
    card = ttk.Frame(root, style="Card.TFrame", padding=24)
    card.pack(fill="both", expand=True, padx=18, pady=18)
    card.columnconfigure(0, weight=1)

    selected = tk.StringVar()
    status = tk.StringVar(value="Выберите заполненный реестр Банка России.")
    saved = None

    def browse():
        path = filedialog.askopenfilename(parent=root, title="Выберите реестр займов МФО",
                                         filetypes=[("Книга Excel", "*.xlsx")])
        if path:
            selected.set(path)
            entry.xview_moveto(1)
            status.set("Файл выбран. Нажмите «Сохранить CSV».")
            action.state(["!disabled"])
            show.state(["disabled"])

    def export():
        nonlocal saved
        source = Path(selected.get())
        name = filedialog.asksaveasfilename(parent=root, title="Сохранить CSV для МоКоД",
                                          initialdir=source.parent, initialfile=source.stem + "_МоКоД.csv",
                                          defaultextension=".csv", filetypes=[("Файл CSV", "*.csv")])
        if not name:
            return
        status.set("Конвертирую реестр…")
        root.update_idletasks()
        show.state(["disabled"])
        try:
            target = Path(name)
            count = convert(source, target)
        except ConversionError as error:
            status.set(str(error))
        except PermissionError:
            status.set("Нет доступа к файлу. Закройте открытый CSV или выберите другую папку.")
        except (BadZipFile, InvalidFileException, OSError):
            status.set("Не удалось прочитать XLSX или записать CSV. Проверьте файл и папку назначения.")
        else:
            saved = target
            status.set(f"CSV сохранён. Строк в выгрузке — {count}.")
            show.state(["!disabled"])

    def show_folder():
        if saved:
            try:
                open_folder(saved.parent)
            except OSError:
                status.set(f"CSV сохранён в папке {saved.parent}")

    ttk.Label(card, text="Реестр займов МФО", style="Title.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
    ttk.Label(card, text="Конвертация XLSX в CSV для МоКоД", style="Card.TLabel").grid(
        row=1, column=0, columnspan=2, sticky="w", pady=(2, 18))
    entry = ttk.Entry(card, textvariable=selected, state="readonly", width=40)
    entry.grid(row=2, column=0, sticky="ew", ipady=4, padx=(0, 10))
    ttk.Button(card, text="Выбрать XLSX…", command=browse).grid(row=2, column=1, sticky="ew")
    action = ttk.Button(card, text="Сохранить CSV", command=export, style="Action.TButton", state="disabled")
    action.grid(row=3, column=0, sticky="w", pady=(14, 0))
    show = ttk.Button(card, text="Открыть папку", command=show_folder, state="disabled")
    show.grid(row=3, column=1, sticky="ew", pady=(14, 0))
    ttk.Label(card, textvariable=status, wraplength=510, style="Card.TLabel").grid(
        row=4, column=0, columnspan=2, sticky="w", pady=(16, 0))
    root.mainloop()
