import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path
import threading

from openpyxl import load_workbook, Workbook
from tkinterdnd2 import TkinterDnD, DND_FILES


class ExcelSumApp:
    def __init__(self, root):
        self.root = root

        self.root.title("Sumowanie składników Excel")
        self.root.geometry("1100x750")
        self.root.minsize(850, 600)

        self.files = []
        self.rows = []

        self.create_styles()
        self.create_interface()

    # ==========================================================
    # WYGLĄD
    # ==========================================================

    def create_styles(self):

        style = ttk.Style()

        try:
            style.theme_use("vista")
        except tk.TclError:
            pass

        style.configure(
            "Title.TLabel",
            font=("Segoe UI", 18, "bold")
        )

        style.configure(
            "Big.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=(15, 9)
        )

        style.configure(
            "Treeview",
            font=("Segoe UI", 10),
            rowheight=28
        )

        style.configure(
            "Treeview.Heading",
            font=("Segoe UI", 10, "bold")
        )

    # ==========================================================
    # INTERFEJS
    # ==========================================================

    def create_interface(self):

        main = ttk.Frame(
            self.root,
            padding=18
        )

        main.pack(
            fill="both",
            expand=True
        )

        # ------------------------------------------------------
        # NAGŁÓWEK
        # ------------------------------------------------------

        header = ttk.Frame(main)

        header.pack(
            fill="x",
            pady=(0, 12)
        )

        ttk.Label(
            header,
            text="Sumowanie składników Excel",
            style="Title.TLabel"
        ).pack(side="left")

        ttk.Button(
            header,
            text="Wybierz pliki Excel",
            command=self.choose_files,
            style="Big.TButton"
        ).pack(side="right")

        # ------------------------------------------------------
        # DRAG & DROP
        # ------------------------------------------------------

        self.drop_area = tk.Frame(
            main,
            bd=2,
            relief="groove",
            height=100
        )

        self.drop_area.pack(
            fill="x",
            pady=(0, 12)
        )

        self.drop_area.pack_propagate(False)

        self.drop_label = tk.Label(
            self.drop_area,
            text=(
                "PRZECIĄGNIJ TUTAJ PLIKI EXCEL\n\n"
                "Możesz przeciągnąć kilka plików .xlsx jednocześnie"
            ),
            font=("Segoe UI", 11, "bold"),
            justify="center"
        )

        self.drop_label.pack(
            fill="both",
            expand=True
        )

        # Rejestracja Drag & Drop
        self.drop_area.drop_target_register(DND_FILES)
        self.drop_area.dnd_bind(
            "<<Drop>>",
            self.drop_files
        )

        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind(
            "<<Drop>>",
            self.drop_files
        )

        # ------------------------------------------------------
        # LISTA PLIKÓW
        # ------------------------------------------------------

        files_label = ttk.Label(
            main,
            text="Wczytane pliki:"
        )

        files_label.pack(
            anchor="w"
        )

        files_frame = ttk.Frame(main)

        files_frame.pack(
            fill="x",
            pady=(4, 10)
        )

        self.file_list = tk.Listbox(
            files_frame,
            height=5,
            font=("Segoe UI", 10)
        )

        self.file_list.pack(
            side="left",
            fill="both",
            expand=True
        )

        files_scroll = ttk.Scrollbar(
            files_frame,
            orient="vertical",
            command=self.file_list.yview
        )

        files_scroll.pack(
            side="right",
            fill="y"
        )

        self.file_list.configure(
            yscrollcommand=files_scroll.set
        )

        file_buttons = ttk.Frame(main)

        file_buttons.pack(
            fill="x",
            pady=(0, 10)
        )

        ttk.Button(
            file_buttons,
            text="Usuń zaznaczony plik",
            command=self.remove_selected_file
        ).pack(side="left")

        ttk.Button(
            file_buttons,
            text="Wyczyść listę",
            command=self.clear
        ).pack(side="left", padx=(8, 0))

        ttk.Button(
            file_buttons,
            text="Przetwórz pliki",
            command=self.start_processing,
            style="Big.TButton"
        ).pack(side="right")

        # ------------------------------------------------------
        # STATUS
        # ------------------------------------------------------

        self.status = tk.StringVar(
            value="Dodaj pliki Excel."
        )

        ttk.Label(
            main,
            textvariable=self.status
        ).pack(
            fill="x",
            pady=(0, 8)
        )

        # ------------------------------------------------------
        # TABELA WYNIKÓW
        # ------------------------------------------------------

        table_frame = ttk.Frame(main)

        table_frame.pack(
            fill="both",
            expand=True
        )

        columns = (
            "number",
            "description",
            "quantity",
            "unit"
        )

        self.tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings"
        )

        self.tree.heading(
            "number",
            text="Numer składnika"
        )

        self.tree.heading(
            "description",
            text="Opis składnika"
        )

        self.tree.heading(
            "quantity",
            text="Ilość"
        )

        self.tree.heading(
            "unit",
            text="Jednostka"
        )

        self.tree.column(
            "number",
            width=160,
            anchor="center"
        )

        self.tree.column(
            "description",
            width=520,
            anchor="w"
        )

        self.tree.column(
            "quantity",
            width=140,
            anchor="e"
        )

        self.tree.column(
            "unit",
            width=120,
            anchor="center"
        )

        scroll_y = ttk.Scrollbar(
            table_frame,
            orient="vertical",
            command=self.tree.yview
        )

        scroll_x = ttk.Scrollbar(
            table_frame,
            orient="horizontal",
            command=self.tree.xview
        )

        self.tree.configure(
            yscrollcommand=scroll_y.set,
            xscrollcommand=scroll_x.set
        )

        self.tree.grid(
            row=0,
            column=0,
            sticky="nsew"
        )

        scroll_y.grid(
            row=0,
            column=1,
            sticky="ns"
        )

        scroll_x.grid(
            row=1,
            column=0,
            sticky="ew"
        )

        table_frame.rowconfigure(
            0,
            weight=1
        )

        table_frame.columnconfigure(
            0,
            weight=1
        )

        # ------------------------------------------------------
        # DOLNY PASEK
        # ------------------------------------------------------

        bottom = ttk.Frame(main)

        bottom.pack(
            fill="x",
            pady=(12, 0)
        )

        self.count = tk.StringVar(
            value="Pliki: 0 | Wiersze: 0 | Pozycje: 0"
        )

        ttk.Label(
            bottom,
            textvariable=self.count
        ).pack(side="left")

        self.save_button = ttk.Button(
            bottom,
            text="Zapisz jako Excel",
            command=self.save_file,
            style="Big.TButton",
            state="disabled"
        )

        self.save_button.pack(
            side="right"
        )

    # ==========================================================
    # DRAG & DROP
    # ==========================================================

    def drop_files(self, event):

        try:

            paths = self.root.tk.splitlist(
                event.data
            )

            excel_files = []

            for path in paths:

                path = path.strip("{}")

                if path.lower().endswith(".xlsx"):

                    if path not in self.files:
                        excel_files.append(path)

            if not excel_files:

                messagebox.showwarning(
                    "Brak plików Excel",
                    "Przeciągnij pliki .xlsx."
                )

                return

            for path in excel_files:

                self.files.append(path)

                self.file_list.insert(
                    tk.END,
                    Path(path).name
                )

            self.status.set(
                f"Dodano {len(excel_files)} plików."
            )

        except Exception as exc:

            messagebox.showerror(
                "Błąd",
                str(exc)
            )

    # ==========================================================
    # WYBÓR PLIKÓW
    # ==========================================================

    def choose_files(self):

        paths = filedialog.askopenfilenames(
            title="Wybierz pliki Excel",
            filetypes=[
                ("Pliki Excel", "*.xlsx"),
                ("Wszystkie pliki", "*.*")
            ]
        )

        if not paths:
            return

        added = 0

        for path in paths:

            if path not in self.files:

                self.files.append(path)

                self.file_list.insert(
                    tk.END,
                    Path(path).name
                )

                added += 1

        self.status.set(
            f"Dodano {added} plików."
        )

    # ==========================================================
    # USUWANIE PLIKU
    # ==========================================================

    def remove_selected_file(self):

        selected = self.file_list.curselection()

        if not selected:
            return

        index = selected[0]

        self.file_list.delete(index)

        del self.files[index]

        self.status.set(
            f"Plików na liście: {len(self.files)}"
        )

    # ==========================================================
    # CZYSZCZENIE
    # ==========================================================

    def clear(self):

        self.files = []
        self.rows = []

        self.file_list.delete(
            0,
            tk.END
        )

        for item in self.tree.get_children():

            self.tree.delete(item)

        self.count.set(
            "Pliki: 0 | Wiersze: 0 | Pozycje: 0"
        )

        self.status.set(
            "Lista wyczyszczona."
        )

        self.save_button.config(
            state="disabled"
        )

    # ==========================================================
    # START PRZETWARZANIA
    # ==========================================================

    def start_processing(self):

        if not self.files:

            messagebox.showwarning(
                "Brak plików",
                "Dodaj przynajmniej jeden plik Excel."
            )

            return

        self.status.set(
            f"Przetwarzanie {len(self.files)} plików..."
        )

        self.save_button.config(
            state="disabled"
        )

        threading.Thread(
            target=self.process_files,
            daemon=True
        ).start()

    # ==========================================================
    # KONWERSJA ILOŚCI
    # ==========================================================

    @staticmethod
    def to_decimal(value):

        if value is None:
            raise InvalidOperation

        if isinstance(value, bool):
            raise InvalidOperation

        if isinstance(
            value,
            (int, float, Decimal)
        ):

            return Decimal(
                str(value)
            )

        text = str(value).strip()

        text = text.replace(
            "\u00a0",
            ""
        )

        text = text.replace(
            " ",
            ""
        )

        text = text.replace(
            ",",
            "."
        )

        return Decimal(text)

    # ==========================================================
    # ROZPOZNAWANIE NAGŁÓWKA
    # ==========================================================

    @staticmethod
    def is_header(value, column_type):

        if value is None:
            return False

        text = (
            str(value)
            .strip()
            .lower()
            .replace("\n", " ")
        )

        if column_type == "number":

            return (
                "numer składnika" in text
                or "numer skladnika" in text
                or "nr składnika" in text
                or "nr skladnika" in text
            )

        if column_type == "description":

            return (
                "opis składnika" in text
                or "opis skladnika" in text
            )

        if column_type == "quantity":

            return (
                "ilość" in text
                or "ilosc" in text
            )

        if column_type == "unit":

            return (
                "jednostka" in text
                or "jm skł" in text
                or "jm skl" in text
            )

        return False

    # ==========================================================
    # USTALANIE PIERWSZEGO WIERSZA DANYCH
    # ==========================================================

    def get_start_row(self, worksheet):

        headers = [
            worksheet.cell(
                1,
                column
            ).value

            for column in (
                2,
                4,
                5,
                6
            )
        ]

        matches = [

            self.is_header(
                headers[0],
                "number"
            ),

            self.is_header(
                headers[1],
                "description"
            ),

            self.is_header(
                headers[2],
                "quantity"
            ),

            self.is_header(
                headers[3],
                "unit"
            )
        ]

        # Najważniejsze:
        # E1 może być:
        # "Ilość"
        # "Ilość skł. (JM skł.)"
        # itd.

        if matches[2]:
            return 2

        if sum(matches) >= 2:
            return 2

        return 1

    # ==========================================================
    # PRZETWARZANIE
    # ==========================================================

    def process_files(self):

        grouped = defaultdict(
            Decimal
        )

        errors = []

        total_rows = 0

        for file_path in self.files:

            try:

                workbook = load_workbook(
                    file_path,
                    read_only=True,
                    data_only=True
                )

                worksheet = workbook.active

                start_row = self.get_start_row(
                    worksheet
                )

                for row_number, row in enumerate(

                    worksheet.iter_rows(
                        min_row=start_row,
                        min_col=2,
                        max_col=6,
                        values_only=True
                    ),

                    start=start_row
                ):

                    # B = row[0]
                    # C = row[1]
                    # D = row[2]
                    # E = row[3]
                    # F = row[4]

                    number = row[0]

                    description = row[2]

                    quantity = row[3]

                    unit = row[4]

                    # Pomijamy pusty wiersz

                    if all(

                        value is None
                        or str(value).strip() == ""

                        for value in (
                            number,
                            description,
                            quantity,
                            unit
                        )
                    ):

                        continue

                    total_rows += 1

                    # Brak numeru składnika

                    if (
                        number is None
                        or str(number).strip() == ""
                    ):

                        errors.append(
                            f"{Path(file_path).name}, "
                            f"wiersz {row_number}: "
                            f"brak Numeru składnika."
                        )

                        continue

                    # Ilość

                    try:

                        qty = self.to_decimal(
                            quantity
                        )

                    except (
                        InvalidOperation,
                        ValueError
                    ):

                        errors.append(
                            f"{Path(file_path).name}, "
                            f"wiersz {row_number}: "
                            f"'{quantity}' w kolumnie E "
                            f"nie jest liczbą."
                        )

                        continue

                    number = str(
                        number
                    ).strip()

                    description = (
                        ""
                        if description is None
                        else str(
                            description
                        ).strip()
                    )

                    unit = (
                        ""
                        if unit is None
                        else str(
                            unit
                        ).strip()
                    )

                    # Grupowanie:
                    #
                    # Numer + Opis + Jednostka

                    key = (
                        number,
                        description,
                        unit
                    )

                    grouped[key] += qty

                workbook.close()

            except Exception as exc:

                errors.append(
                    f"{Path(file_path).name}: "
                    f"nie można odczytać pliku — "
                    f"{exc}"
                )

        # ------------------------------------------------------
        # WYNIK
        # ------------------------------------------------------

        result = []

        for key, total in grouped.items():

            result.append(
                (
                    key[0],
                    key[1],
                    total,
                    key[2]
                )
            )

        result.sort(
            key=lambda x: x[0]
        )

        self.root.after(
            0,
            self.show_result,
            result,
            errors,
            total_rows
        )

    # ==========================================================
    # WYŚWIETLENIE WYNIKU
    # ==========================================================

    @staticmethod
    def format_number(value):

        if value == value.to_integral():

            return str(
                int(value)
            )

        text = format(
            value.normalize(),
            "f"
        )

        if "." in text:

            text = (
                text
                .rstrip("0")
                .rstrip(".")
            )

        return text

    def show_result(
        self,
        result,
        errors,
        total_rows
    ):

        self.rows = result

        # Czyszczenie tabeli

        for item in self.tree.get_children():

            self.tree.delete(item)

        # Wynik

        for (
            number,
            description,
            quantity,
            unit
        ) in result:

            self.tree.insert(
                "",
                "end",
                values=(
                    number,
                    description,
                    self.format_number(
                        quantity
                    ),
                    unit
                )
            )

        self.count.set(
            f"Pliki: {len(self.files)} | "
            f"Wiersze: {total_rows} | "
            f"Pozycje po zsumowaniu: "
            f"{len(result)}"
        )

        if result:

            self.save_button.config(
                state="normal"
            )

        if errors:

            self.status.set(
                f"Gotowe. Znaleziono "
                f"{len(errors)} problemów."
            )

            preview = "\n".join(
                errors[:20]
            )

            if len(errors) > 20:

                preview += (
                    f"\n... oraz "
                    f"{len(errors) - 20} kolejnych."
                )

            messagebox.showwarning(
                "Problemy w danych",
                "Pliki zostały przetworzone, "
                "ale znaleziono problemy:\n\n"
                + preview
            )

        else:

            self.status.set(
                "Gotowe. Wszystkie pliki "
                "przetworzono poprawnie."
            )

    # ==========================================================
    # ZAPIS EXCEL
    # ==========================================================

    def save_file(self):

        if not self.rows:

            messagebox.showwarning(
                "Brak danych",
                "Nie ma danych do zapisania."
            )

            return

        path = filedialog.asksaveasfilename(

            title="Zapisz wynik",

            defaultextension=".xlsx",

            initialfile=(
                "zsumowane_skladniki.xlsx"
            ),

            filetypes=[
                (
                    "Plik Excel",
                    "*.xlsx"
                )
            ]
        )

        if not path:
            return

        try:

            workbook = Workbook()

            worksheet = workbook.active

            worksheet.title = (
                "Zsumowane składniki"
            )

            # Nagłówki

            worksheet.append(
                [
                    "Numer składnika",
                    "Opis składnika",
                    "Ilość",
                    "Jednostka"
                ]
            )

            # Dane

            for (
                number,
                description,
                quantity,
                unit
            ) in self.rows:

                worksheet.append(
                    [
                        number,
                        description,
                        float(quantity),
                        unit
                    ]
                )

            worksheet.freeze_panes = "A2"

            worksheet.auto_filter.ref = (
                worksheet.dimensions
            )

            worksheet.column_dimensions[
                "A"
            ].width = 20

            worksheet.column_dimensions[
                "B"
            ].width = 60

            worksheet.column_dimensions[
                "C"
            ].width = 15

            worksheet.column_dimensions[
                "D"
            ].width = 15

            workbook.save(
                path
            )

            self.status.set(
                "Zapisano wynik."
            )

            messagebox.showinfo(
                "Zapisano",
                "Gotowy plik zapisano jako:\n\n"
                + path
            )

        except PermissionError:

            messagebox.showerror(
                "Brak dostępu",
                "Nie można zapisać pliku. "
                "Sprawdź, czy plik nie jest "
                "otwarty w Excelu."
            )

        except Exception as exc:

            messagebox.showerror(
                "Błąd zapisu",
                str(exc)
            )


# ==============================================================
# START PROGRAMU
# ==============================================================

def main():

    root = TkinterDnD.Tk()

    ExcelSumApp(
        root
    )

    root.mainloop()


if __name__ == "__main__":

    main()
