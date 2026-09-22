"""Tests for the shared bordered-block renderer behind Table.show and grid.show."""

from __future__ import annotations

from pyhandlexl._display import render_box


class TestNoRows:
    def test_empty_input_is_an_empty_string(self):
        assert render_box([]) == ""


class TestBorders:
    def test_top_and_bottom_rule_match(self):
        lines = render_box([["a", "b"], ["c", "d"]]).splitlines()
        assert lines[0] == lines[-1]
        assert lines[0].startswith("+") and lines[0].endswith("+")

    def test_content_lines_are_piped(self):
        lines = render_box([["a"]]).splitlines()
        assert lines[1].startswith("|") and lines[1].endswith("|")

    def test_a_header_separator_appears_after_header_rows(self):
        lines = render_box([["h"], ["1"], ["2"]], header_rows=1).splitlines()
        rules = [line for line in lines if line.startswith("+")]
        assert len(rules) == 3  # top, after the header, bottom
        assert lines[0] == lines[2] == lines[5]  # every rule is identical

    def test_no_doubled_rule_when_the_header_is_the_only_row(self):
        lines = render_box([["h"]], header_rows=1).splitlines()
        assert len(lines) == 3  # top rule, header, bottom rule — not two rules in a row
        rules = [line for line in lines if line.startswith("+")]
        assert len(rules) == 2

    def test_no_separator_without_header_rows(self):
        lines = render_box([["a"], ["b"]]).splitlines()
        rules = [line for line in lines if line.startswith("+")]
        assert len(rules) == 2


class TestAlignment:
    def test_columns_are_padded_to_their_widest_value(self):
        lines = render_box([["x", "long value"], ["a", "y"]]).splitlines()
        widths = {len(line) for line in lines}
        assert len(widths) == 1  # every line is the same length


class TestValueFormatting:
    def test_none_is_blank(self):
        lines = render_box([[None]]).splitlines()
        assert lines[1] == "|  |"

    def test_non_strings_are_stringified(self):
        assert "3.5" in render_box([[3.5]]) and "True" in render_box([[True]])

    def test_a_newline_is_made_visible_not_a_real_line_break(self):
        text = render_box([["a\nb"]])
        assert text.count("\n") == 2  # only the two rule/content separators, not a third
        assert "a\\nb" in text

    def test_a_carriage_return_is_made_visible_too(self):
        assert "a\\nb" in render_box([["a\rb"]])
        assert "a\\nb" in render_box([["a\r\nb"]])


class TestRealisticShape:
    def test_a_table_like_block_with_truncation(self):
        rows = [["", "v"], ["r0", "0"], ["...", "..."], ["r9", "9"]]
        text = render_box(rows, header_rows=1)
        lines = text.splitlines()
        assert len(lines) == 7  # top, header, sep, r0, ..., r9, bottom
        assert "..." in lines[4]
