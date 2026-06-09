"""Графический интерфейс (Tkinter) — для пользователей без консоли.

Выбор файла + настройки (поля, номер страницы, полужирность, распознавание структуры)
→ оформление по ГОСТ 7.32-2017. Сверх stdlib зависимостей не добавляет (tkinter встроен).

Логика форматирования вынесена в format_file() — её можно тестировать без UI.
(mypy для этого модуля отключён в pyproject: tkinter-стабы слишком строги к UI-глю.)
"""

from __future__ import annotations

import sys
from pathlib import Path

from .formatter import format_document
from .profile import build_profile

# Подпись в выпадающем списке → значение для профиля.
PAGE_NUMBER_LABELS: dict[str, str] = {
    "Внизу по центру": "bottom-center",
    "Внизу справа": "bottom-right",
    "Вверху справа": "top-right",
    "Без нумерации": "none",
}


def default_output(input_path: str) -> str:
    """INPUT.docx → INPUT.gost.docx (рядом, исходник не трогаем)."""
    p = Path(input_path)
    return str(p.parent / (p.stem + ".gost.docx"))


def format_file(
    input_path: str,
    output_path: str | None = None,
    *,
    detect_structure: bool = True,
    margins: str | None = None,
    page_number: str | None = None,
    page_number_size_pt: int | None = None,
    bold_headings: bool | None = None,
) -> tuple[str, list[str]]:
    """Собрать профиль из опций и отформатировать. Вернуть (путь результата, сообщения).

    Поднимает GostDocError (плохой формат/файл) и ValueError (плохие поля) — UI их ловит.
    """
    profile = build_profile(
        margins=margins,
        page_number=page_number,
        page_number_size_pt=page_number_size_pt,
        bold_headings=bold_headings,
    )
    out = output_path or default_output(input_path)
    messages = format_document(input_path, out, detect_structure=detect_structure, profile=profile)
    return out, messages


def main(argv: list[str] | None = None) -> int:
    """Запустить окно. Если передан путь к файлу (перетаскивание на .exe) — подставить его."""
    import tkinter as tk  # импорт внутри — чтобы CLI/тесты не требовали дисплея

    args = sys.argv[1:] if argv is None else argv
    initial_file = args[0] if args and not args[0].startswith("-") else None

    root = tk.Tk()
    _App(root, initial_file)
    root.mainloop()
    return 0


class _App:
    def __init__(self, root, initial_file: str | None = None) -> None:
        import tkinter as tk
        from tkinter import filedialog, scrolledtext, ttk

        self._tk = tk
        self._ttk = ttk
        self._filedialog = filedialog
        self.root = root
        root.title("gostdoc — оформление по ГОСТ 7.32-2017")
        root.geometry("640x560")
        root.minsize(560, 480)

        pad = {"padx": 10, "pady": 4}

        # --- Файл ---
        file_frame = ttk.LabelFrame(root, text="Документ")
        file_frame.pack(fill="x", **pad)
        self.file_var = tk.StringVar(value=initial_file or "")
        ttk.Entry(file_frame, textvariable=self.file_var).pack(
            side="left", fill="x", expand=True, padx=8, pady=8
        )
        ttk.Button(file_frame, text="Выбрать .docx…", command=self._choose_file).pack(
            side="right", padx=8, pady=8
        )

        # --- Настройки ---
        opt = ttk.LabelFrame(root, text="Настройки оформления")
        opt.pack(fill="x", **pad)

        self.detect_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            opt, text="Распознавать заголовки, главы и подписи", variable=self.detect_var
        ).grid(row=0, column=0, columnspan=4, sticky="w", padx=8, pady=4)

        self.bold_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt, text="Полужирные заголовки", variable=self.bold_var).grid(
            row=1, column=0, columnspan=4, sticky="w", padx=8, pady=4
        )

        ttk.Label(opt, text="Поля (мм):  Л").grid(row=2, column=0, sticky="e", padx=(8, 0))
        self.ml = self._num_entry(opt, "30", row=2, col=1)
        ttk.Label(opt, text="П").grid(row=2, column=2, sticky="e")
        self.mr = self._num_entry(opt, "15", row=2, col=3)
        ttk.Label(opt, text="В").grid(row=3, column=0, sticky="e", padx=(8, 0))
        self.mt = self._num_entry(opt, "20", row=3, col=1)
        ttk.Label(opt, text="Н").grid(row=3, column=2, sticky="e")
        self.mb = self._num_entry(opt, "20", row=3, col=3)

        ttk.Label(opt, text="Номер страницы:").grid(row=4, column=0, sticky="e", padx=(8, 0), pady=4)
        self.page_var = tk.StringVar(value="Внизу по центру")
        ttk.Combobox(
            opt, textvariable=self.page_var, values=list(PAGE_NUMBER_LABELS), state="readonly", width=18
        ).grid(row=4, column=1, columnspan=2, sticky="w", pady=4)

        ttk.Label(opt, text="Кегль №:").grid(row=4, column=3, sticky="e")
        self.size_var = tk.StringVar(value="14")
        ttk.Entry(opt, textvariable=self.size_var, width=5).grid(row=4, column=4, sticky="w", padx=4)

        # --- Кнопка ---
        self.run_btn = ttk.Button(root, text="Оформить по ГОСТ", command=self._run)
        self.run_btn.pack(fill="x", **pad)

        # --- Лог ---
        log_frame = ttk.LabelFrame(root, text="Результат")
        log_frame.pack(fill="both", expand=True, **pad)
        self.log = scrolledtext.ScrolledText(log_frame, height=12, state="disabled", wrap="word")
        self.log.pack(fill="both", expand=True, padx=8, pady=8)

        self._say("Выберите файл .docx и нажмите «Оформить по ГОСТ».")
        self._say("Результат сохранится рядом: ВашФайл.gost.docx (исходник не меняется).")

    def _num_entry(self, parent, default: str, *, row: int, col: int):
        var = self._tk.StringVar(value=default)
        self._ttk.Entry(parent, textvariable=var, width=5).grid(
            row=row, column=col, sticky="w", padx=4
        )
        return var

    def _choose_file(self) -> None:
        path = self._filedialog.askopenfilename(
            title="Выберите документ .docx", filetypes=[("Документы Word", "*.docx"), ("Все файлы", "*.*")]
        )
        if path:
            self.file_var.set(path)

    def _say(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")
        self.root.update_idletasks()

    def _run(self) -> None:
        from tkinter import messagebox

        from .formatter import GostDocError

        path = self.file_var.get().strip()
        if not path:
            messagebox.showwarning("Нет файла", "Сначала выберите файл .docx.")
            return

        margins = f"{self.ml.get()},{self.mr.get()},{self.mt.get()},{self.mb.get()}"
        try:
            size = int(self.size_var.get())
        except ValueError:
            messagebox.showerror("Ошибка", "Кегль номера страницы должен быть числом.")
            return

        self.run_btn.configure(state="disabled")
        self._say("\nОбрабатываю…")
        try:
            out, messages = format_file(
                path,
                detect_structure=self.detect_var.get(),
                margins=margins,
                page_number=PAGE_NUMBER_LABELS[self.page_var.get()],
                page_number_size_pt=size,
                bold_headings=self.bold_var.get(),
            )
        except (GostDocError, ValueError) as exc:
            self._say(f"Ошибка: {exc}")
            messagebox.showerror("Ошибка", str(exc))
            return
        except Exception as exc:  # noqa: BLE001 — не дать окну «упасть» молча
            self._say(f"Непредвиденная ошибка: {exc}")
            messagebox.showerror("Ошибка", str(exc))
            return
        finally:
            self.run_btn.configure(state="normal")

        self._say(f"Готово: {out}")
        for m in messages:
            self._say("  " + m if not m.startswith(" ") else m)
        messagebox.showinfo("Готово", f"Файл оформлен:\n{out}")


if __name__ == "__main__":
    sys.exit(main())
