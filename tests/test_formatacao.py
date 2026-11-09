import codecs

import pytest
from openpyxl import Workbook, load_workbook

from formatacao import (
    FormatError,
    FormatOptions,
    ensure_distinct_paths,
    format_table,
    list_sheets,
    read_table,
    write_table,
)


class TestFormatTable:
    def test_trims_and_collapses_spaces_in_text_cells(self):
        table = [["  Nome ", "Cidade"], ["  Ana   Maria ", "São Paulo"]]

        result = format_table(table, FormatOptions(remove_empty_rows=False))

        assert result == [["Nome", "Cidade"], ["Ana Maria", "São Paulo"]]

    def test_keeps_spaces_when_trim_is_disabled(self):
        table = [["a  b"], ["  c "]]

        result = format_table(table, FormatOptions(trim_spaces=False))

        assert result == [["a  b"], ["  c "]]

    def test_removes_empty_rows_but_keeps_header(self):
        table = [["a", "b"], ["", None], ["1", "2"], ["  ", ""]]

        assert format_table(table, FormatOptions()) == [["a", "b"], ["1", "2"]]

    def test_keeps_empty_rows_when_disabled(self):
        table = [["a", "b"], ["", None], ["1", "2"]]

        result = format_table(table, FormatOptions(remove_empty_rows=False))

        assert result == [["a", "b"], ["", None], ["1", "2"]]

    def test_removes_only_columns_where_every_cell_is_empty(self):
        table = [["a", "", "c"], ["1", "", ""], ["2", "", "x"]]

        result = format_table(table, FormatOptions(remove_empty_columns=True))

        assert result == [["a", "c"], ["1", ""], ["2", "x"]]

    def test_removes_duplicate_rows_keeping_first_occurrence(self):
        table = [["a"], ["x"], ["y"], ["x"]]

        result = format_table(table, FormatOptions(remove_duplicate_rows=True))

        assert result == [["a"], ["x"], ["y"]]

    def test_header_is_not_considered_when_removing_duplicates(self):
        table = [["x"], ["x"]]

        result = format_table(table, FormatOptions(remove_duplicate_rows=True))

        assert result == [["x"], ["x"]]

    @pytest.mark.parametrize(
        ("mode", "expected"),
        [
            ("upper", "NOME COMPLETO"),
            ("lower", "nome completo"),
            ("title", "Nome Completo"),
        ],
    )
    def test_changes_header_case_only(self, mode, expected):
        table = [["nome completo"], ["Ana"]]

        result = format_table(table, FormatOptions(header_case=mode))

        assert result == [[expected], ["Ana"]]

    def test_pads_ragged_rows_to_the_same_width(self):
        table = [["a", "b", "c"], ["1"], ["2", "3"]]

        result = format_table(table, FormatOptions(remove_empty_rows=False))

        assert result == [["a", "b", "c"], ["1", None, None], ["2", "3", None]]

    def test_empty_table_returns_empty_table(self):
        assert format_table([], FormatOptions()) == []

    def test_does_not_modify_input_table(self):
        table = [["  a  "], ["", ], ["  b "]]
        original = [list(row) for row in table]

        format_table(table, FormatOptions())

        assert table == original

    def test_rejects_unknown_header_case(self):
        with pytest.raises(ValueError):
            FormatOptions(header_case="invalido")


class TestCsv:
    def test_detects_semicolon_delimiter(self, tmp_path):
        path = tmp_path / "dados.csv"
        path.write_text("nome;valor\nAna;10\n", encoding="utf-8")

        assert read_table(path) == [["nome", "valor"], ["Ana", "10"]]

    def test_reads_cp1252_encoded_file(self, tmp_path):
        path = tmp_path / "dados.csv"
        path.write_bytes("nome,cidade\nJosé,São Paulo\n".encode("cp1252"))

        assert read_table(path)[1] == ["José", "São Paulo"]

    def test_explicit_delimiter_overrides_detection(self, tmp_path):
        path = tmp_path / "dados.csv"
        path.write_text("a|b\n1|2\n", encoding="utf-8")

        assert read_table(path, delimiter="|") == [["a", "b"], ["1", "2"]]

    def test_writes_utf8_bom_and_chosen_delimiter(self, tmp_path):
        output = tmp_path / "saida.csv"

        write_table(output, [["nome", "valor"], ["José", None]], FormatOptions(csv_delimiter=";"))

        raw = output.read_bytes()
        assert raw.startswith(codecs.BOM_UTF8)
        assert raw.decode("utf-8-sig").splitlines() == ["nome;valor", "José;"]

    def test_converts_csv_to_excel(self, tmp_path):
        source = tmp_path / "entrada.csv"
        source.write_text("nome,valor\nAna,10\n", encoding="utf-8")
        output = tmp_path / "saida.xlsx"

        write_table(output, read_table(source), FormatOptions())

        values = [[cell.value for cell in row] for row in load_workbook(output).active.iter_rows()]
        assert values == [["nome", "valor"], ["Ana", "10"]]


class TestExcel:
    def test_roundtrip_keeps_values_and_applies_style(self, tmp_path):
        output = tmp_path / "saida.xlsx"
        table = [["Nome", "Valor"], ["Ana", 10], ["Bruno", None]]

        write_table(output, table, FormatOptions(excel_style=True))

        sheet = load_workbook(output).active
        assert [[cell.value for cell in row] for row in sheet.iter_rows()] == table
        assert sheet["A1"].font.bold is True
        assert sheet.freeze_panes == "A2"
        assert sheet.auto_filter.ref == "A1:B3"

    def test_skips_style_when_disabled(self, tmp_path):
        output = tmp_path / "saida.xlsx"

        write_table(output, [["Nome"], ["Ana"]], FormatOptions(excel_style=False))

        sheet = load_workbook(output).active
        assert sheet["A1"].font.bold is not True
        assert sheet.freeze_panes is None

    def test_text_starting_with_equals_is_written_as_text(self, tmp_path):
        output = tmp_path / "saida.xlsx"

        write_table(output, [["f"], ["=1+1"]], FormatOptions(excel_style=False))

        cell = load_workbook(output).active["A2"]
        assert cell.value == "=1+1"
        assert cell.data_type == "s"

    def test_lists_sheets_and_reads_named_sheet(self, tmp_path):
        path = tmp_path / "planilha.xlsx"
        workbook = Workbook()
        first = workbook.active
        first.title = "Clientes"
        first.append(["a"])
        second = workbook.create_sheet("Vendas")
        second.append(["b"])
        workbook.save(path)

        assert list_sheets(path) == ["Clientes", "Vendas"]
        assert read_table(path, sheet="Vendas") == [["b"]]

    def test_unknown_sheet_raises_format_error(self, tmp_path):
        path = tmp_path / "planilha.xlsx"
        Workbook().save(path)

        with pytest.raises(FormatError):
            read_table(path, sheet="Nao existe")

    def test_invalid_excel_file_raises_format_error(self, tmp_path):
        path = tmp_path / "corrompido.xlsx"
        path.write_bytes(b"isto nao e um xlsx")

        with pytest.raises(FormatError):
            read_table(path)


class TestErrors:
    def test_unsupported_input_extension_raises(self, tmp_path):
        path = tmp_path / "dados.txt"
        path.write_text("a,b", encoding="utf-8")

        with pytest.raises(FormatError):
            read_table(path)

    def test_unsupported_output_extension_raises(self, tmp_path):
        with pytest.raises(FormatError):
            write_table(tmp_path / "saida.xlsm", [["a"]], FormatOptions())

    def test_blocks_saving_over_the_input_file(self, tmp_path):
        path = tmp_path / "dados.csv"
        path.write_text("a\n", encoding="utf-8")

        with pytest.raises(FormatError):
            ensure_distinct_paths(path, path)

    def test_allows_different_output_path(self, tmp_path):
        ensure_distinct_paths(tmp_path / "a.csv", tmp_path / "a_formatado.csv")
