"""Regras de leitura, formatação e gravação de planilhas CSV e Excel.

Este módulo não depende de interface gráfica, para que as regras possam ser
testadas isoladamente.
"""

from __future__ import annotations

import csv
import io
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, List, Optional

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.worksheet import Worksheet

Row = List[Any]
Table = List[Row]

CSV_EXTENSIONS = (".csv",)
EXCEL_EXTENSIONS = (".xlsx", ".xlsm")
OUTPUT_EXTENSIONS = (".csv", ".xlsx")
HEADER_CASES = ("none", "upper", "lower", "title")

_SNIFF_DELIMITERS = ",;\t|"
_SNIFF_SAMPLE_SIZE = 4096
_HORIZONTAL_SPACES = re.compile(r"[ \t\xa0]+")
_HEADER_FONT = Font(bold=True)
_HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
_THIN_SIDE = Side(style="thin", color="BFBFBF")
_CELL_BORDER = Border(left=_THIN_SIDE, right=_THIN_SIDE, top=_THIN_SIDE, bottom=_THIN_SIDE)
_HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)
_HEADER_ROW_HEIGHT = 24
_MIN_COLUMN_WIDTH = 8
_MAX_COLUMN_WIDTH = 60
# Brazilian format: "1.234,56", "1234", "-0,5". Dots are only valid as thousands separators.
_BR_NUMBER_RE = re.compile(r"^-?(0|[1-9]\d{0,2}(\.\d{3})+|[1-9]\d*)(,\d+)?$")
# Above 15 significant digits Excel stores numbers inexactly, so those stay as text.
_MAX_EXACT_DIGITS = 15
_INTEGER_FORMAT = "#,##0"


class FormatError(Exception):
    """Erro esperado ao ler ou gravar uma planilha. A mensagem é mostrada ao usuário."""


@dataclass(frozen=True)
class FormatOptions:
    trim_spaces: bool = True
    remove_empty_rows: bool = True
    remove_empty_columns: bool = False
    remove_duplicate_rows: bool = False
    header_case: str = "none"
    csv_delimiter: str = ","
    excel_style: bool = True
    # Only affects Excel output. Off by default, because codes such as CEP or CPF
    # would lose their meaning if they looked like numbers.
    convert_numbers: bool = False

    def __post_init__(self) -> None:
        if self.header_case not in HEADER_CASES:
            raise ValueError(f"header_case inválido: {self.header_case!r}")


def read_table(
    path: str | Path,
    sheet: Optional[str] = None,
    delimiter: Optional[str] = None,
) -> Table:
    """Read a CSV or Excel file into a list of rows. The first row is the header.

    `sheet` is used only for Excel files and `delimiter` only for CSV files.
    When `delimiter` is None, it is detected automatically.
    """
    source = Path(path)
    extension = source.suffix.lower()
    if extension in EXCEL_EXTENSIONS:
        return _read_excel(source, sheet)
    if extension in CSV_EXTENSIONS:
        return _read_csv(source, delimiter)
    raise _unsupported_format(source)


def list_sheets(path: str | Path) -> List[str]:
    workbook = _open_workbook(Path(path))
    try:
        return list(workbook.sheetnames)
    finally:
        workbook.close()


def format_table(table: Table, options: FormatOptions) -> Table:
    """Return a new table with the requested rules applied. The input is not changed."""
    if not table:
        return []

    width = max(len(row) for row in table)
    padded = [list(row) + [None] * (width - len(row)) for row in table]

    if options.trim_spaces:
        padded = [[_clean_text(value) for value in row] for row in padded]

    header = padded[0]
    body = padded[1:]

    if options.header_case != "none":
        header = [_change_case(value, options.header_case) for value in header]
    if options.remove_empty_rows:
        body = [row for row in body if not _is_empty_row(row)]
    if options.remove_duplicate_rows:
        body = _unique_rows(body)

    result = [header] + body
    if options.remove_empty_columns:
        result = _drop_empty_columns(result)
    return result


def write_table(path: str | Path, table: Table, options: FormatOptions) -> None:
    """Write a table as CSV or XLSX, chosen by the file extension."""
    target = Path(path)
    extension = target.suffix.lower()
    if extension == ".csv":
        _write_csv(target, table, options.csv_delimiter)
    elif extension == ".xlsx":
        _write_excel(target, table, options)
    else:
        raise FormatError(
            f"Formato de saída não suportado: '{target.suffix or target.name}'. "
            "Use .csv ou .xlsx."
        )


def ensure_distinct_paths(input_path: str | Path, output_path: str | Path) -> None:
    """Block saving over the source file, which would destroy the original data."""
    if Path(input_path).resolve() == Path(output_path).resolve():
        raise FormatError("O arquivo de saída não pode ser o mesmo arquivo de entrada.")


def _read_excel(path: Path, sheet_name: Optional[str]) -> Table:
    workbook = _open_workbook(path)
    try:
        worksheet = _get_sheet(workbook, sheet_name)
        return [list(row) for row in worksheet.iter_rows(values_only=True)]
    finally:
        workbook.close()


def _open_workbook(path: Path):
    try:
        return load_workbook(path, read_only=True, data_only=True)
    except (OSError, InvalidFileException, zipfile.BadZipFile, KeyError, ValueError) as exc:
        raise FormatError(f"Não foi possível abrir a planilha '{path.name}'.") from exc


def _get_sheet(workbook, sheet_name: Optional[str]) -> Worksheet:
    if sheet_name is None:
        return workbook.active
    if sheet_name not in workbook.sheetnames:
        raise FormatError(f"A aba '{sheet_name}' não existe nesta planilha.")
    return workbook[sheet_name]


def _read_csv(path: Path, delimiter: Optional[str]) -> Table:
    try:
        raw_bytes = path.read_bytes()
    except OSError as exc:
        raise FormatError(f"Não foi possível ler o arquivo '{path.name}'.") from exc

    text = _decode(raw_bytes)
    separator = delimiter or _sniff_delimiter(text)
    try:
        return list(csv.reader(io.StringIO(text, newline=""), delimiter=separator))
    except csv.Error as exc:
        raise FormatError(f"O arquivo CSV está malformado: {exc}") from exc


def _decode(raw_bytes: bytes) -> str:
    # utf-8-sig also strips the BOM that Excel adds to CSV files.
    for encoding in ("utf-8-sig", "cp1252"):
        try:
            return raw_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw_bytes.decode("latin-1")


def _sniff_delimiter(text: str) -> str:
    try:
        sample = text[:_SNIFF_SAMPLE_SIZE]
        return csv.Sniffer().sniff(sample, delimiters=_SNIFF_DELIMITERS).delimiter
    except csv.Error:
        return ","


def _write_csv(path: Path, table: Table, delimiter: str) -> None:
    try:
        with path.open("w", encoding="utf-8-sig", newline="") as handle:
            csv.writer(handle, delimiter=delimiter).writerows(table)
    except OSError as exc:
        raise FormatError(f"Não foi possível salvar '{path.name}'.") from exc


def _write_excel(path: Path, table: Table, options: FormatOptions) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Dados"

    for row_index, row in enumerate(table, start=1):
        for column_index, value in enumerate(row, start=1):
            cell_value = value
            number_format = None
            if row_index > 1 and options.convert_numbers and isinstance(value, str):
                parsed = _parse_br_number(value)
                if parsed is not value:
                    cell_value = parsed
                    number_format = _number_format(value, parsed)

            cell = sheet.cell(row=row_index, column=column_index, value=cell_value)
            if isinstance(cell_value, str) and cell_value.startswith("="):
                # Text starting with "=" must stay text, not become a formula.
                cell.data_type = "s"
            if number_format is not None:
                cell.number_format = number_format

    if options.excel_style and table:
        _style_sheet(sheet)

    try:
        workbook.save(path)
    except OSError as exc:
        raise FormatError(f"Não foi possível salvar '{path.name}'.") from exc


def _style_sheet(sheet: Worksheet) -> None:
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, max_col=sheet.max_column):
        for cell in row:
            cell.border = _CELL_BORDER

    for cell in sheet[1]:
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = _HEADER_ALIGNMENT
    sheet.row_dimensions[1].height = _HEADER_ROW_HEIGHT

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    for column in sheet.iter_cols(min_row=1, max_row=sheet.max_row):
        longest = max(len(_as_text(cell.value)) for cell in column)
        width = min(max(longest + 2, _MIN_COLUMN_WIDTH), _MAX_COLUMN_WIDTH)
        sheet.column_dimensions[column[0].column_letter].width = width


def _clean_text(value: Any) -> Any:
    if isinstance(value, str):
        return _HORIZONTAL_SPACES.sub(" ", value).strip()
    return value


def _change_case(value: Any, mode: str) -> Any:
    if not isinstance(value, str):
        return value
    if mode == "upper":
        return value.upper()
    if mode == "lower":
        return value.lower()
    return value.title()


def _is_empty(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _is_empty_row(row: Row) -> bool:
    return all(_is_empty(value) for value in row)


def _unique_rows(rows: Table) -> Table:
    seen = set()
    unique: Table = []
    for row in rows:
        key = tuple(row)
        if key not in seen:
            seen.add(key)
            unique.append(row)
    return unique


def _drop_empty_columns(table: Table) -> Table:
    width = len(table[0])
    kept = [
        index
        for index in range(width)
        if not all(_is_empty(row[index]) for row in table)
    ]
    return [[row[index] for index in kept] for row in table]


def _as_text(value: Any) -> str:
    return "" if value is None else str(value)


def _parse_br_number(value: Any) -> Any:
    """Return the number for a Brazilian-formatted string, or the value unchanged."""
    if not isinstance(value, str):
        return value

    text = value.strip()
    if not _BR_NUMBER_RE.match(text):
        return value
    # "1.234" without a comma is ambiguous, so it stays as text.
    if "." in text and "," not in text:
        return value

    normalized = text.replace(".", "").replace(",", ".")
    digits = normalized.replace("-", "").replace(".", "")
    if len(digits.lstrip("0")) > _MAX_EXACT_DIGITS:
        return value

    if "." in normalized:
        return float(normalized)
    return int(normalized)


def _number_format(text: str, number: Any) -> str:
    """Keep the decimal places written in the source text, e.g. "1.234,50" -> 2 places."""
    if isinstance(number, int):
        return _INTEGER_FORMAT
    decimals = len(text.strip().rsplit(",", 1)[1])
    return f"#,##0.{'0' * decimals}"


def _unsupported_format(path: Path) -> FormatError:
    return FormatError(
        f"Formato não suportado: '{path.suffix or path.name}'. "
        "Use arquivos .csv, .xlsx ou .xlsm."
    )
