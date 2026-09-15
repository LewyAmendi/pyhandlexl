"""Tests for the TableStyle dataclass: validation, presets, equality."""

from __future__ import annotations

import pytest

from pyhandlexl.style import TableStyle


class TestDefaults:
    def test_default_values(self):
        s = TableStyle()
        assert s.header_font_name == "Calibri"
        assert s.header_font_size == 11
        assert s.header_font_color == "000000"
        assert s.header_bold is True
        assert s.header_fill == "76933C"
        assert s.data_font_name == "Calibri"
        assert s.data_font_size == 11
        assert s.data_font_color == "000000"
        assert s.band_fill == "F2F2F2"
        assert s.border_color == "000000"

    def test_is_frozen(self):
        s = TableStyle()
        with pytest.raises(AttributeError):
            s.header_bold = False

    def test_equality(self):
        assert TableStyle() == TableStyle()
        assert TableStyle(header_bold=False) != TableStyle()


class TestPresets:
    def test_default_preset_matches_plain_construction(self):
        assert TableStyle() == TableStyle.DEFAULT

    def test_minimal_has_no_fills(self):
        assert TableStyle.MINIMAL.header_fill == ""
        assert TableStyle.MINIMAL.band_fill == ""
        assert TableStyle.MINIMAL.header_bold is True
        assert TableStyle.MINIMAL.border_color != ""

    def test_none_has_nothing(self):
        assert TableStyle.NONE.header_fill == ""
        assert TableStyle.NONE.band_fill == ""
        assert TableStyle.NONE.header_bold is False
        assert TableStyle.NONE.border_color == ""

    def test_presets_are_distinct(self):
        assert TableStyle.DEFAULT != TableStyle.MINIMAL
        assert TableStyle.MINIMAL != TableStyle.NONE
        assert TableStyle.DEFAULT != TableStyle.NONE


class TestValidation:
    def test_empty_header_font_name_raises(self):
        with pytest.raises(ValueError):
            TableStyle(header_font_name="")

    def test_non_string_font_name_raises(self):
        with pytest.raises(ValueError):
            TableStyle(data_font_name=123)

    def test_zero_font_size_raises(self):
        with pytest.raises(ValueError):
            TableStyle(header_font_size=0)

    def test_negative_font_size_raises(self):
        with pytest.raises(ValueError):
            TableStyle(data_font_size=-1)

    def test_non_int_font_size_raises(self):
        with pytest.raises(ValueError):
            TableStyle(header_font_size="11")

    def test_bool_font_size_raises(self):
        # bool is technically an int subclass in Python — explicitly rejected
        with pytest.raises(ValueError):
            TableStyle(header_font_size=True)

    def test_non_bool_header_bold_raises(self):
        with pytest.raises(TypeError):
            TableStyle(header_bold="yes")

    def test_invalid_header_font_color_raises(self):
        with pytest.raises(ValueError):
            TableStyle(header_font_color="not-a-color")

    def test_empty_header_font_color_raises(self):
        # unlike fill/border colors, font colors are always required
        with pytest.raises(ValueError):
            TableStyle(header_font_color="")

    def test_short_hex_color_raises(self):
        with pytest.raises(ValueError):
            TableStyle(data_font_color="FFF")

    def test_empty_header_fill_is_allowed(self):
        TableStyle(header_fill="")  # no raise — means "no fill"

    def test_empty_band_fill_is_allowed(self):
        TableStyle(band_fill="")

    def test_empty_border_color_is_allowed(self):
        TableStyle(border_color="")

    def test_invalid_header_fill_raises(self):
        with pytest.raises(ValueError):
            TableStyle(header_fill="nope")
