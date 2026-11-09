"""Interface gráfica para formatar arquivos CSV e Excel."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Optional

from formatacao import (
    CSV_EXTENSIONS,
    EXCEL_EXTENSIONS,
    OUTPUT_EXTENSIONS,
    FormatError,
    FormatOptions,
    Table,
    ensure_distinct_paths,
    format_table,
    list_sheets,
    read_table,
    write_table,
)

# Drag and drop needs tkinterdnd2. Without it, the app still works with the "Procurar..." button.
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    _TK_BASE = TkinterDnD.Tk
except ImportError:
    DND_FILES = None
    _TK_BASE = tk.Tk

PREVIEW_LIMIT = 100
AUTO_DELIMITER = "Automático"
INPUT_DELIMITERS = {
    AUTO_DELIMITER: None,
    "Vírgula (,)": ",",
    "Ponto e vírgula (;)": ";",
    "Tabulação": "\t",
    "Barra vertical (|)": "|",
}
OUTPUT_DELIMITERS = {
    "Vírgula (,)": ",",
    "Ponto e vírgula (;)": ";",
    "Tabulação": "\t",
    "Barra vertical (|)": "|",
}
HEADER_CASES = {
    "Não alterar": "none",
    "MAIÚSCULAS": "upper",
    "minúsculas": "lower",
    "Título (Primeira Letra Maiúscula)": "title",
}
INPUT_FILE_TYPES = [
    ("Planilhas", "*.csv *.xlsx *.xlsm"),
    ("CSV", "*.csv"),
    ("Excel", "*.xlsx *.xlsm"),
    ("Todos os arquivos", "*.*"),
]
OUTPUT_FILE_TYPES = [
    ("Excel", "*.xlsx"),
    ("CSV", "*.csv"),
]


def _is_excel(path: str) -> bool:
    return Path(path).suffix.lower() in EXCEL_EXTENSIONS


def _suggest_output_path(input_path: str) -> str:
    source = Path(input_path)
    extension = ".xlsx" if _is_excel(input_path) else ".csv"
    return str(source.with_name(f"{source.stem}_formatado{extension}"))


def _as_text(value: object) -> str:
    return "" if value is None else str(value)


def _data_rows(table: Table) -> int:
    return max(len(table) - 1, 0)


def _width(table: Table) -> int:
    return max((len(row) for row in table), default=0)


class FormatterApp(_TK_BASE):
    def __init__(self) -> None:
        super().__init__()
        self.title("Formatador de Planilhas")
        self.geometry("960x720")
        self.minsize(800, 600)

        self._table: Table = []

        self.input_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.input_delimiter_var = tk.StringVar(value=AUTO_DELIMITER)
        self.trim_var = tk.BooleanVar(value=True)
        self.remove_empty_rows_var = tk.BooleanVar(value=True)
        self.remove_empty_columns_var = tk.BooleanVar(value=False)
        self.remove_duplicates_var = tk.BooleanVar(value=False)
        self.header_case_var = tk.StringVar(value="Não alterar")
        self.output_var = tk.StringVar()
        self.output_delimiter_var = tk.StringVar(value="Vírgula (,)")
        self.excel_style_var = tk.BooleanVar(value=True)
        self.convert_numbers_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Arraste um CSV ou Excel para esta janela ou clique em Procurar.")

        self._build_ui()
        self.output_var.trace_add("write", lambda *_: self._update_output_options())
        self._update_input_options()
        self._update_output_options()
        self._enable_drop()

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)

        self._build_input_section().grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 8))
        self._build_rules_section().grid(row=1, column=0, sticky="ew", padx=16, pady=8)
        self._build_output_section().grid(row=2, column=0, sticky="ew", padx=16, pady=8)
        self._build_preview_section().grid(row=3, column=0, sticky="nsew", padx=16, pady=8)
        self._build_footer().grid(row=4, column=0, sticky="ew", padx=16, pady=(8, 16))

    def _build_input_section(self) -> ttk.LabelFrame:
        section = ttk.LabelFrame(self, text="1. Arquivo de entrada", padding=12)
        section.columnconfigure(1, weight=1)

        ttk.Label(section, text="Arquivo:").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(section, textvariable=self.input_var, state="readonly").grid(
            row=0, column=1, sticky="ew"
        )
        ttk.Button(section, text="Procurar...", command=self._choose_input).grid(
            row=0, column=2, padx=(8, 0)
        )

        ttk.Label(section, text="Aba (Excel):").grid(
            row=1, column=0, sticky="w", padx=(0, 8), pady=(8, 0)
        )
        self.sheet_combo = ttk.Combobox(
            section, textvariable=self.sheet_var, state="disabled", width=30
        )
        self.sheet_combo.grid(row=1, column=1, sticky="w", pady=(8, 0))
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda _event: self._load_source())

        ttk.Label(section, text="Separador (CSV):").grid(
            row=2, column=0, sticky="w", padx=(0, 8), pady=(8, 0)
        )
        self.input_delimiter_combo = ttk.Combobox(
            section,
            textvariable=self.input_delimiter_var,
            values=list(INPUT_DELIMITERS),
            state="readonly",
            width=30,
        )
        self.input_delimiter_combo.grid(row=2, column=1, sticky="w", pady=(8, 0))
        self.input_delimiter_combo.bind("<<ComboboxSelected>>", lambda _event: self._load_source())
        return section

    def _build_rules_section(self) -> ttk.LabelFrame:
        section = ttk.LabelFrame(self, text="2. Regras de formatação", padding=12)
        section.columnconfigure(0, weight=1)
        section.columnconfigure(1, weight=1)

        rules = [
            (self.trim_var, "Remover espaços extras (início, fim e entre palavras)"),
            (self.remove_empty_rows_var, "Remover linhas totalmente vazias"),
            (self.remove_empty_columns_var, "Remover colunas totalmente vazias"),
            (self.remove_duplicates_var, "Remover linhas duplicadas (mantém a primeira)"),
        ]
        for index, (variable, label) in enumerate(rules):
            ttk.Checkbutton(
                section, text=label, variable=variable, command=self._render_preview
            ).grid(row=index // 2, column=index % 2, sticky="w", padx=(0, 16), pady=2)

        ttk.Label(section, text="Maiúsculas/minúsculas do cabeçalho:").grid(
            row=2, column=0, sticky="w", pady=(8, 0)
        )
        header_combo = ttk.Combobox(
            section,
            textvariable=self.header_case_var,
            values=list(HEADER_CASES),
            state="readonly",
            width=32,
        )
        header_combo.grid(row=2, column=1, sticky="w", pady=(8, 0))
        header_combo.bind("<<ComboboxSelected>>", lambda _event: self._render_preview())
        return section

    def _build_output_section(self) -> ttk.LabelFrame:
        section = ttk.LabelFrame(self, text="3. Arquivo de saída", padding=12)
        section.columnconfigure(1, weight=1)

        ttk.Label(section, text="Salvar em:").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(section, textvariable=self.output_var, state="readonly").grid(
            row=0, column=1, sticky="ew"
        )
        ttk.Button(section, text="Escolher...", command=self._choose_output).grid(
            row=0, column=2, padx=(8, 0)
        )

        ttk.Label(section, text="Separador (CSV de saída):").grid(
            row=1, column=0, sticky="w", padx=(0, 8), pady=(8, 0)
        )
        self.output_delimiter_combo = ttk.Combobox(
            section,
            textvariable=self.output_delimiter_var,
            values=list(OUTPUT_DELIMITERS),
            state="readonly",
            width=30,
        )
        self.output_delimiter_combo.grid(row=1, column=1, sticky="w", pady=(8, 0))

        self.excel_style_check = ttk.Checkbutton(
            section,
            text="Estilizar planilha Excel (cabeçalho em negrito, filtros, cabeçalho fixo e largura automática)",
            variable=self.excel_style_var,
        )
        self.excel_style_check.grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))

        self.convert_numbers_check = ttk.Checkbutton(
            section,
            text="Converter números brasileiros (1.234,56) em números do Excel. Não use para CEP, CPF ou telefone.",
            variable=self.convert_numbers_var,
        )
        self.convert_numbers_check.grid(row=3, column=0, columnspan=3, sticky="w", pady=(4, 0))
        return section

    def _build_preview_section(self) -> ttk.LabelFrame:
        section = ttk.LabelFrame(
            self, text=f"Pré-visualização (até {PREVIEW_LIMIT} linhas)", padding=12
        )
        section.columnconfigure(0, weight=1)
        section.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(section, show="headings")
        vertical_scroll = ttk.Scrollbar(section, orient="vertical", command=self.tree.yview)
        horizontal_scroll = ttk.Scrollbar(section, orient="horizontal", command=self.tree.xview)
        self.tree.configure(
            yscrollcommand=vertical_scroll.set,
            xscrollcommand=horizontal_scroll.set,
        )

        self.tree.grid(row=0, column=0, sticky="nsew")
        vertical_scroll.grid(row=0, column=1, sticky="ns")
        horizontal_scroll.grid(row=1, column=0, sticky="ew")
        return section

    def _build_footer(self) -> ttk.Frame:
        footer = ttk.Frame(self)
        footer.columnconfigure(0, weight=1)

        ttk.Label(footer, textvariable=self.status_var, wraplength=600).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Button(footer, text="Formatar e salvar", command=self._save).grid(
            row=0, column=1, padx=(8, 0)
        )
        return footer

    def _choose_input(self) -> None:
        path = filedialog.askopenfilename(
            title="Escolha a planilha", filetypes=INPUT_FILE_TYPES
        )
        if path:
            self._set_input(path)

    def _enable_drop(self) -> None:
        if DND_FILES is None:
            return
        try:
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._on_drop)
        except tk.TclError:
            # tkdnd could not be loaded in this Tk build; the file button still works.
            pass

    def _on_drop(self, event) -> str:
        dropped = self.tk.splitlist(event.data)
        if not dropped:
            return event.action

        path = dropped[0]
        if Path(path).suffix.lower() not in (*CSV_EXTENSIONS, *EXCEL_EXTENSIONS):
            messagebox.showwarning(
                "Arquivo não suportado",
                "Arraste um arquivo .csv, .xlsx ou .xlsm.",
            )
            return event.action

        self._set_input(path)
        return event.action

    def _set_input(self, path: str) -> None:
        self.input_var.set(path)
        self.output_var.set(_suggest_output_path(path))
        self._update_input_options()
        self._refresh_sheets()
        self._load_source()

    def _choose_output(self) -> None:
        current = Path(self.output_var.get()) if self.output_var.get() else None
        path = filedialog.asksaveasfilename(
            title="Salvar arquivo formatado como",
            initialdir=str(current.parent) if current else "",
            initialfile=current.name if current else "",
            defaultextension=".xlsx",
            filetypes=OUTPUT_FILE_TYPES,
        )
        if not path:
            return
        if Path(path).suffix.lower() not in OUTPUT_EXTENSIONS:
            path += ".xlsx"
        self.output_var.set(path)

    def _update_input_options(self) -> None:
        is_excel = _is_excel(self.input_var.get())
        self.sheet_combo.configure(state="readonly" if is_excel else "disabled")
        self.input_delimiter_combo.configure(state="disabled" if is_excel else "readonly")

    def _update_output_options(self) -> None:
        suffix = Path(self.output_var.get()).suffix.lower()
        self.output_delimiter_combo.configure(
            state="readonly" if suffix == ".csv" else "disabled"
        )
        excel_state = "normal" if suffix == ".xlsx" else "disabled"
        self.excel_style_check.configure(state=excel_state)
        self.convert_numbers_check.configure(state=excel_state)

    def _refresh_sheets(self) -> None:
        path = self.input_var.get()
        sheets: list[str] = []
        if _is_excel(path):
            try:
                sheets = list_sheets(path)
            except FormatError as exc:
                messagebox.showerror("Não foi possível ler a planilha", str(exc))
        self.sheet_combo.configure(values=sheets)
        self.sheet_var.set(sheets[0] if sheets else "")

    def _selected_sheet(self) -> Optional[str]:
        return self.sheet_var.get() or None

    def _selected_input_delimiter(self) -> Optional[str]:
        return INPUT_DELIMITERS[self.input_delimiter_var.get()]

    def _load_source(self) -> None:
        path = self.input_var.get()
        if not path:
            return
        try:
            self._table = read_table(
                path,
                sheet=self._selected_sheet(),
                delimiter=self._selected_input_delimiter(),
            )
        except FormatError as exc:
            self._table = []
            messagebox.showerror("Não foi possível ler o arquivo", str(exc))
        self._render_preview()

    def _build_options(self) -> FormatOptions:
        return FormatOptions(
            trim_spaces=self.trim_var.get(),
            remove_empty_rows=self.remove_empty_rows_var.get(),
            remove_empty_columns=self.remove_empty_columns_var.get(),
            remove_duplicate_rows=self.remove_duplicates_var.get(),
            header_case=HEADER_CASES[self.header_case_var.get()],
            csv_delimiter=OUTPUT_DELIMITERS[self.output_delimiter_var.get()],
            excel_style=self.excel_style_var.get(),
            convert_numbers=self.convert_numbers_var.get(),
        )

    def _render_preview(self) -> None:
        options = self._build_options()
        formatted = format_table(self._table, options)

        self.tree.delete(*self.tree.get_children())
        width = _width(formatted)
        columns = [f"col{index}" for index in range(width)]
        self.tree.configure(columns=columns)

        if formatted:
            for index, column_id in enumerate(columns):
                title = _as_text(formatted[0][index]) or f"Coluna {index + 1}"
                self.tree.heading(column_id, text=title, anchor="w")
                self.tree.column(column_id, width=150, anchor="w", stretch=True)
            for row in formatted[1 : PREVIEW_LIMIT + 1]:
                self.tree.insert("", "end", values=[_as_text(value) for value in row])

        self._update_summary(formatted)

    def _update_summary(self, formatted: Table) -> None:
        if not self._table:
            return
        self.status_var.set(
            f"Original: {_data_rows(self._table)} linhas × {_width(self._table)} colunas"
            f"  →  Formatado: {_data_rows(formatted)} linhas × {_width(formatted)} colunas"
        )

    def _save(self) -> None:
        output = self.output_var.get()
        if not self._table:
            messagebox.showwarning(
                "Nada para salvar", "Escolha um arquivo de entrada válido primeiro."
            )
            return
        if not output:
            messagebox.showwarning("Nada para salvar", "Escolha onde salvar o arquivo.")
            return

        try:
            ensure_distinct_paths(self.input_var.get(), output)
            # Default to "No" so that a dismissed dialog never overwrites a file.
            if Path(output).exists() and not messagebox.askyesno(
                "Substituir arquivo",
                f"O arquivo já existe. Substituir?\n\n{output}",
                default=messagebox.NO,
            ):
                return
            options = self._build_options()
            formatted = format_table(self._table, options)
            write_table(output, formatted, options)
        except FormatError as exc:
            messagebox.showerror("Erro ao salvar", str(exc))
            return

        self.status_var.set(f"Arquivo salvo: {output}")
        messagebox.showinfo("Concluído", f"Arquivo salvo em:\n{output}")


def main() -> None:
    app = FormatterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
