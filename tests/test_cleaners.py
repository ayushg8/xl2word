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
