import pytest
from openpyxl import load_workbook

from formatacao import FormatOptions, _parse_br_number, write_table


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("1.234,56", 1234.56),
        ("1234,5", 1234.5),
        ("0,5", 0.5),
        ("-3,2", -3.2),
        ("10", 10),
        ("1.234.567,89", 1234567.89),
    ],
)
def test_parses_brazilian_numbers(text, expected):
    assert _parse_br_number(text) == expected


@pytest.mark.parametrize(
    "text",
    [
        "007",  # leading zero: looks like a code
        "01310100",  # CEP
        "123.456.789-09",  # CPF
        "11987654321.0",  # dot is not a thousands separator
        "1.234",  # dot without comma is ambiguous
        "",
        "abc",
        "1,2,3",
        "12345678901234567890,5",  # more digits than Excel stores exactly
    ],
)
def test_keeps_text_that_is_not_a_brazilian_number(text):
    assert _parse_br_number(text) == text


def test_non_string_values_are_returned_unchanged():
    assert _parse_br_number(3.5) == 3.5
    assert _parse_br_number(None) is None


def test_converts_numbers_in_excel_body_only(tmp_path):
    output = tmp_path / "saida.xlsx"
    table = [["1,5", "CEP"], ["1.234,56", "01310100"], ["10", "texto"]]

    write_table(output, table, FormatOptions(convert_numbers=True, excel_style=False))

    sheet = load_workbook(output).active
    assert sheet["A1"].value == "1,5"  # header stays text
    assert sheet["A2"].value == 1234.56
    assert sheet["A2"].number_format == "#,##0.00"
    assert sheet["A3"].value == 10
    assert sheet["A3"].number_format == "#,##0"
    assert sheet["B2"].value == "01310100"  # CEP stays text


def test_number_conversion_is_off_by_default(tmp_path):
    output = tmp_path / "saida.xlsx"

    write_table(output, [["valor"], ["1.234,56"]], FormatOptions(excel_style=False))

    assert load_workbook(output).active["A2"].value == "1.234,56"


def test_keeps_decimal_places_from_source_text(tmp_path):
    output = tmp_path / "saida.xlsx"

    write_table(output, [["v"], ["1,123"]], FormatOptions(convert_numbers=True, excel_style=False))

    assert load_workbook(output).active["A2"].number_format == "#,##0.000"


def test_excel_style_adds_borders_and_header_layout(tmp_path):
    output = tmp_path / "saida.xlsx"

    write_table(output, [["Nome", "Valor"], ["Ana", 10]], FormatOptions(excel_style=True))

    sheet = load_workbook(output).active
    assert sheet["A2"].border.left.style == "thin"
    assert sheet["B2"].border.bottom.style == "thin"
    assert sheet["A1"].alignment.horizontal == "center"
    assert sheet["A1"].alignment.wrap_text is True
    assert sheet.row_dimensions[1].height == 24
