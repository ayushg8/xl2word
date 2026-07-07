import datetime as dt
from xl2word.cleaners import format_cell_value

def test_none_is_empty():
    assert format_cell_value(None, "0.00") == ""

def test_general_passthrough():
    assert format_cell_value("Loading", None) == "Loading"
    assert format_cell_value("Loading", "General") == "Loading"

def test_decimals():
    assert format_cell_value(25.4, "0.00") == "25.40"
    assert format_cell_value(1.087, "0.000") == "1.087"

def test_integer():
    assert format_cell_value(142, "0") == "142"

def test_percent():
    assert format_cell_value(0.012, "0.0%") == "1.2%"
    assert format_cell_value(0.5, "0%") == "50%"

def test_thousands():
    assert format_cell_value(1190, "#,##0") == "1,190"

def test_date():
    assert format_cell_value(dt.datetime(2026, 4, 1), "yyyy-mm-dd") == "2026-04-01"

def test_fallback():
    assert format_cell_value(3.14159, "weird") == "3.14159"

def test_scientific_notation():
    # was silently collapsing to 0.00 (data loss)
    assert format_cell_value(1.85e-6, "0.00E+00") == "1.85E-06"

def test_quoted_unit_literals():
    assert format_cell_value(128, '0" °C"') == "128 °C"
    assert format_cell_value(2.52, '0.00" g/cm³"') == "2.52 g/cm³"
    assert format_cell_value(5000, '#,##0" MΩ"') == "5,000 MΩ"

def test_currency_and_accounting():
    assert format_cell_value(12.47, '"$"#,##0.00') == "$12.47"
    assert format_cell_value(-12.40, '#,##0.00;(#,##0.00)') == "(12.40)"

def test_datetime_with_time():
    assert format_cell_value(dt.datetime(2026, 6, 10, 14, 30), "yyyy-mm-dd hh:mm") == "2026-06-10 14:30"

def test_time_of_day():
    assert format_cell_value(dt.time(12, 0), "h:mm") == "12:00"

def test_error_strings_hidden():
    for e in ("#REF!", "#DIV/0!", "#VALUE!", "#N/A"):
        assert format_cell_value(e, "General") == ""
