"""
tests/test_helpers.py — Phase 3
Unit tests for pure helper functions that don't require Excel.
Tests: injection escaping, color parsing, type inference, validation helpers.
"""

import pytest
import sys
import os

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from shared_context import (
    escape_injection, escape_data,
    validate_range_address, validate_sheet_name,
    check_data_size, MAX_CELLS,
)
from format_tools import hex_to_rgb, rgb_to_bgr_int
from table_tools import _detect_format, _col_num_to_letter
from verification import _values_equal, _col_num_to_letter as v_col_letter
from security import redact_secrets, validate_path, is_macro_enabled


# ── Injection Guard Tests ─────────────────────────────────────────────────────

class TestInjectionEscaping:
    def test_formula_escaped(self):
        assert escape_injection("=SUM(A1:A10)") == "'=SUM(A1:A10)"

    def test_plus_prefix_escaped(self):
        assert escape_injection("+123") == "'+123"

    def test_minus_prefix_escaped(self):
        assert escape_injection("-abc") == "'-abc"

    def test_at_prefix_escaped(self):
        assert escape_injection("@user") == "'@user"

    def test_normal_string_unchanged(self):
        assert escape_injection("Hello World") == "Hello World"

    def test_number_unchanged(self):
        assert escape_injection(42) == 42
        assert escape_injection(3.14) == 3.14

    def test_none_unchanged(self):
        assert escape_injection(None) is None

    def test_empty_string_unchanged(self):
        assert escape_injection("") == ""

    def test_escape_data_list_of_lists(self):
        data = [["Name", "=EVIL()"], ["Alice", "+DROP TABLE"]]
        result = escape_data(data)
        assert result[0][1] == "'=EVIL()"
        assert result[1][1] == "'+DROP TABLE"
        assert result[0][0] == "Name"

    def test_escape_data_list_of_dicts(self):
        data = [{"Name": "Alice", "Formula": "=WEBSERVICE(url)"}]
        result = escape_data(data)
        assert result[0]["Formula"] == "'=WEBSERVICE(url)"


# ── Color Parsing Tests ───────────────────────────────────────────────────────

class TestColorParsing:
    def test_hex_to_rgb_normal(self):
        assert hex_to_rgb("#FF0000") == (255, 0, 0)
        assert hex_to_rgb("#00FF00") == (0, 255, 0)
        assert hex_to_rgb("#1B365D") == (27, 54, 93)

    def test_hex_to_rgb_no_hash(self):
        assert hex_to_rgb("FFFFFF") == (255, 255, 255)

    def test_hex_to_rgb_shorthand(self):
        assert hex_to_rgb("#FFF") == (255, 255, 255)
        assert hex_to_rgb("#000") == (0, 0, 0)

    def test_hex_to_rgb_invalid(self):
        with pytest.raises(ValueError):
            hex_to_rgb("#GGGGGG")

    def test_hex_to_rgb_none(self):
        assert hex_to_rgb(None) is None

    def test_rgb_to_bgr_int(self):
        # White: R=255, G=255, B=255 → 255 + 255*256 + 255*65536 = 16777215
        assert rgb_to_bgr_int((255, 255, 255)) == 16777215
        # Red: R=255, G=0, B=0 → 255
        assert rgb_to_bgr_int((255, 0, 0)) == 255
        # Blue: R=0, G=0, B=255 → 255*65536 = 16711680
        assert rgb_to_bgr_int((0, 0, 255)) == 16711680


# ── Column Number/Letter Conversion ──────────────────────────────────────────

class TestColumnConversion:
    def test_col_num_to_letter(self):
        assert _col_num_to_letter(1) == "A"
        assert _col_num_to_letter(26) == "Z"
        assert _col_num_to_letter(27) == "AA"
        assert _col_num_to_letter(52) == "AZ"
        assert _col_num_to_letter(53) == "BA"
        assert _col_num_to_letter(702) == "ZZ"
        assert _col_num_to_letter(703) == "AAA"


# ── Number Format Detection ───────────────────────────────────────────────────

class TestFormatDetection:
    def test_date_detection(self):
        from datetime import date
        values = [date(2024, 1, 1), date(2024, 2, 1), date(2024, 3, 1)]
        assert _detect_format(values) == "yyyy-mm-dd"

    def test_currency_detection(self):
        values = [12345.67, 98765.43, 55000.0]
        fmt = _detect_format(values)
        assert "#,##0" in fmt or fmt == "General"

    def test_percentage_detection(self):
        values = [0.12, 0.45, 0.98, 0.3]
        assert _detect_format(values) == "0.00%"

    def test_text_no_format(self):
        values = ["Alice", "Bob", "Charlie"]
        assert _detect_format(values) == ""

    def test_empty_no_format(self):
        assert _detect_format([]) == ""
        assert _detect_format([None, None]) == ""


# ── Input Validation ──────────────────────────────────────────────────────────

class TestInputValidation:
    def test_validate_range_address_normal(self):
        assert validate_range_address("A1") == "A1"
        assert validate_range_address("a1:z100") == "A1:Z100"

    def test_validate_range_address_empty(self):
        with pytest.raises(ValueError):
            validate_range_address("")

    def test_validate_range_address_none(self):
        with pytest.raises(ValueError):
            validate_range_address(None)

    def test_validate_sheet_name_valid(self):
        assert validate_sheet_name("Sheet1") == "Sheet1"
        assert validate_sheet_name("  My Sheet  ") == "My Sheet"

    def test_validate_sheet_name_too_long(self):
        with pytest.raises(ValueError):
            validate_sheet_name("A" * 32)

    def test_validate_sheet_name_forbidden_chars(self):
        for char in r'\/:*?[]':
            with pytest.raises(ValueError):
                validate_sheet_name(f"Sheet{char}1")

    def test_check_data_size_ok(self):
        data = [["A", "B"]] * 100
        check_data_size(data)  # Should not raise

    def test_check_data_size_exceeds(self):
        # 1000 rows × 600 cols = 600,000 > MAX_CELLS
        data = [["x"] * 600] * 1001
        with pytest.raises(ValueError, match="exceeds the limit"):
            check_data_size(data)


# ── Values Equality ───────────────────────────────────────────────────────────

class TestValuesEqual:
    def test_none_equality(self):
        assert _values_equal(None, None) is True

    def test_numeric_within_tolerance(self):
        assert _values_equal(1.0, 1.0000000001) is True

    def test_numeric_outside_tolerance(self):
        assert _values_equal(1.0, 2.0) is False

    def test_string_equality(self):
        assert _values_equal("Hello", "Hello") is True
        assert _values_equal("Hello", "World") is False

    def test_string_with_whitespace(self):
        assert _values_equal("  Hello  ", "Hello") is True


# ── Security Tests ────────────────────────────────────────────────────────────

class TestSecurity:
    def test_redact_secrets(self):
        args = {"sheet": "Sheet1", "password": "mysecret", "token": "abc123", "data": "value"}
        redacted = redact_secrets(args)
        assert redacted["password"] == "***REDACTED***"
        assert redacted["token"] == "***REDACTED***"
        assert redacted["sheet"] == "Sheet1"
        assert redacted["data"] == "value"

    def test_is_macro_enabled(self):
        assert is_macro_enabled("file.xlsm") is True
        assert is_macro_enabled("file.xlsx") is False
        assert is_macro_enabled("file.xlam") is True

    def test_path_traversal_detected(self):
        with pytest.raises(ValueError, match="traversal"):
            validate_path("../../Windows/System32/evil.exe")


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
