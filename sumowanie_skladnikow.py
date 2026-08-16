import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path
import threading

from openpyxl import load_workbook, Workbook


class ExcelSumApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sumowanie składników Excel — wiele plików")
        self.root.geometry("1050x680")
        self.root.minsize(800, 520)

        self.rows = []
        self.files = []

        self.setup_style()
        self.build_ui()

    def setup_style(self):
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
            padding=(14, 8)
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

    def build_ui(self):
        outer = ttk.Frame(self.root, padding=18)
        outer.pack(fill="both", expand=True)

        # Nagłówek
        header = ttk.Frame(outer)
        header.pack(fill="x", pady=(0, 12))

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

        # Informacja
        info = tk.Frame(
            outer,
            bd=1,
            relief="solid",
            padx=15,
            pady=14
        )
        info.pack(fill="x", pady=(0, 10))

        tk.Label(
            info,
            text="MOŻESZ WYBRAĆ WIELE PLIKÓW .XLSX",
            font=("Segoe UI", 12, "bold")
        ).pack()

        tk.Label(
            info,
            text=(
                "Zaznacz kilka plików jednocześnie. "
                "Program połączy je i zsumuje wspólne składniki."
            ),
            font=("Segoe UI", 10)
        ).pack(pady=(5, 0))

        # Status
        self.status_var = tk.StringVar(
            value="Gotowe. Wybierz jeden lub więcej plików Excel."
        )

        ttk.Label(
            outer,
            textvariable=self.status_var
        ).pack(fill="x", pady=(0, 8))

        # Tabela
        table_frame = ttk.Frame(outer)
        table_frame.pack(fill="both", expand=True)

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

        scrollbar_y = ttk.Scrollbar(
            table_frame,
            orient="vertical",
            command=self.tree.yview
        )

        scrollbar_x = ttk.Scrollbar(
            table_frame,
            orient="horizontal",
            command=self.tree.xview
        )

        self.tree.configure(
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set
        )

        self.tree.grid(
            row=0,
            column=0,
            sticky="nsew"
        )

        scrollbar_y.grid(
            row=0,
            column=1,
            sticky="ns"
        )

        scrollbar_x.grid(
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

        # Dolny pasek
        bottom = ttk.Frame(outer)
        bottom.pack(
            fill="x",
            pady=(12, 0)
        )

        self.count_var = tk.StringVar(
            value="Pliki: 0 | Wiersze: 0 | Pozycje: 0"
        )

        ttk.Label(
            bottom,
            textvariable=self.count_var
        ).pack(side="left")

        ttk.Button(
            bottom,
            text="Wyczyść",
            command=self.clear
        ).pack(
            side="right",
            padx=(8, 0)
        )

        self.save_button = ttk.Button(
            bottom,
            text="Zapisz jako Excel",
            command=self.save_file,
            style="Big.TButton",
            state="disabled"
        )

        self.save_button.pack(side="right")

    # ---------------------------------------------------------
    # WYBÓR WIELU PLIKÓW
    # ---------------------------------------------------------

    def choose_files(self):
        paths = filedialog.askopenfilenames(
            title="Wybierz pliki Excel",
            filetypes=[
                ("Pliki Excel", "*.xlsx"),
                ("Wszystkie pliki", "*.*")
            ]
        )

        if paths:
            self.process_files(list(paths))

    # ---------------------------------------------------------
    # CZYSZCZENIE
    # ---------------------------------------------------------

    def clear(self):
        self.files = []
        self.rows = []

        for item in self.tree.get_children():
            self.tree.delete(item)

        self.count_var.set(
            "Pliki: 0 | Wiersze: 0 | Pozycje: 0"
        )

        self.status_var.set(
            "Wyczyszczono. Wybierz pliki Excel."
        )

        self.save_button.config(
            state="disabled"
        )

    # ---------------------------------------------------------
    # PRZETWARZANIE
    # ---------------------------------------------------------

    def process_files(self, paths):
        self.files = paths

        self.status_var.set(
            f"Wczytywanie {len(paths)} plików..."
        )

        self.save_button.config(
            state="disabled"
        )

        thread = threading.Thread(
            target=self.worker,
            args=(paths,),
            daemon=True
        )

        thread.start()

    # ---------------------------------------------------------
    # KONWERSJA ILOŚCI
    # ---------------------------------------------------------

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
            return Decimal(str(value))

        text = str(value).strip()

        # Obsługa np. "12,5"
        text = text.replace(" ", "")
        text = text.replace(",", ".")

        return Decimal(text)

    # ---------------------------------------------------------
    # PRZETWARZANIE PLIKÓW
    # ---------------------------------------------------------

    def worker(self, paths):

        grouped = defaultdict(Decimal)
        errors = []

        total_data_rows = 0

        try:

            for file_path in paths:

                try:

                    workbook = load_workbook(
                        file_path,
                        read_only=True,
                        data_only=True
                    )

                    worksheet = workbook.active

                    # Sprawdzenie nagłówków
                    headers = [
                        worksheet.cell(1, column).value
                        for column in (2, 4, 5, 6)
                    ]

                    expected = [
                        {
                            "Numer składnika",
                            "Numer skladnika",
                            "Nr składnika",
                            "Nr skladnika"
                        },
                        {
                            "Opis składnika",
                            "Opis skladnika"
                        },
                        {
                            "Ilość",
                            "Ilosc"
                        },
                        {
                            "Jednostka",
                            "Jednostka miary"
                        }
                    ]

                    has_header = all(
                        headers[i] is not None
                        and str(headers[i]).strip()
                        in expected[i]
                        for i in range(4)
                    )

                    start_row = 2 if has_header else 1

                    for row_number, row in enumerate(
                        worksheet.iter_rows(
                            min_row=start_row,
                            min_col=2,
                            max_col=6,
                            values_only=True
                        ),
                        start=start_row
                    ):

                        # B,C,D,E,F
                        number = row[0]
                        description = row[2]
                        quantity = row[3]
                        unit = row[4]

                        # Pomijamy całkowicie puste wiersze
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

                        total_data_rows += 1

                        # Numer składnika
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

                            quantity_decimal = self.to_decimal(
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

                        number = str(number).strip()

                        description = (
                            ""
                            if description is None
                            else str(description).strip()
                        )

                        unit = (
                            ""
                            if unit is None
                            else str(unit).strip()
                        )

                        # KLUCZ GRUPOWANIA
                        #
                        # Ten sam:
                        # Numer + Opis + Jednostka
                        #
                        # zostanie zsumowany.

                        key = (
                            number,
                            description,
                            unit
                        )

                        grouped[key] += quantity_decimal

                    workbook.close()

                except Exception as exc:

                    errors.append(
                        f"{Path(file_path).name}: "
                        f"nie można odczytać pliku — {exc}"
                    )

            # Wynik
            result = []

            for key, total in grouped.items():

                number = key[0]
                description = key[1]
                unit = key[2]

                result.append(
                    (
                        number,
                        description,
                        total,
                        unit
                    )
                )

            # Sortowanie po numerze
            result.sort(
                key=lambda row: row[0]
            )

            self.root.after(
                0,
                self.show_result,
                result,
                errors,
                total_data_rows
            )

        except Exception as exc:

            self.root.after(
                0,
                lambda: messagebox.showerror(
                    "Błąd",
                    str(exc)
                )
            )

    # ---------------------------------------------------------
    # WYŚWIETLENIE WYNIKU
    # ---------------------------------------------------------

    @staticmethod
    def format_decimal(value):

        if value == value.to_integral():
            return str(int(value))

        text = format(
            value.normalize(),
            "f"
        )

        if "." in text:
            text = text.rstrip("0").rstrip(".")

        return text

    def show_result(
        self,
        result,
        errors,
        total_data_rows
    ):

        self.rows = result

        # Czyszczenie tabeli
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Wstawianie danych
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
                    self.format_decimal(quantity),
                    unit
                )
            )

        self.count_var.set(
            f"Pliki: {len(self.files)} | "
            f"Wiersze: {total_data_rows} | "
            f"Pozycje po zsumowaniu: {len(result)}"
        )

        if result:
            self.save_button.config(
                state="normal"
            )
        else:
            self.save_button.config(
                state="disabled"
            )

        if errors:

            self.status_var.set(
                f"Gotowe. Pominięto "
                f"{len(errors)} błędnych wierszy."
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
                "Błędy w danych",
                "Pliki zostały przetworzone, "
                "ale znaleziono problemy:\n\n"
                + preview
            )

        else:

            self.status_var.set(
                "Gotowe. Wszystkie pliki "
                "przetworzono poprawnie."
            )

    # ---------------------------------------------------------
    # ZAPIS WYNIKU
    # ---------------------------------------------------------

    def save_file(self):

        if not self.rows:
            messagebox.showinfo(
                "Brak danych",
                "Najpierw wczytaj pliki Excel."
            )

            return

        path = filedialog.asksaveasfilename(
            title="Zapisz wynik",
            defaultextension=".xlsx",
            initialfile="zsumowane_skladniki.xlsx",
            filetypes=[
                ("Plik Excel", "*.xlsx")
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

            workbook.save(path)

            self.status_var.set(
                f"Zapisano wynik: {path}"
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
                "Sprawdź, czy nie jest otwarty w Excelu."
            )

        except Exception as exc:

            messagebox.showerror(
                "Błąd zapisu",
                str(exc)
            )


def main():

    root = tk.Tk()

    ExcelSumApp(root)

    root.mainloop()


if __name__ == "__main__":
    main()
