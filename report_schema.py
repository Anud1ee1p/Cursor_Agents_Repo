from __future__ import annotations

from typing import Any, Callable, NamedTuple, Sequence

# ======================================================================================
# Palette
# ======================================================================================

NAVY = "1F4E78"
SUBTITLE_GRAY = "595959"
GRID_GRAY = "D9D9D9"

# Table status fills — soft tints, readable behind black text in a normal-width table.
FILL_OK_HEX = "C6E0B4"        # green      — check passed
FILL_WARN_HEX = "FFE699"      # amber      — informational / no data / not attempted
FILL_FAIL_HEX = "F8CBAD"      # orange     — failed
FILL_REVIEW_HEX = "9DC3E6"    # blue       — in spec but flagged for review
FILL_SPEC_HEX = "DDEBF7"      # pale blue  — spec reference row
FILL_OUT_OF_TOL_HEX = "FF7B7B"  # red      — single value outside its derived tolerance

# Dense-map fills — the same five meanings at full saturation, for the 72-column maps
# where a cell is 3-5 characters wide and a soft tint would not read at all.
MAP_PASS_HEX = "92D050"       # green      — taught / good
MAP_FAIL_HEX = "FF6D6D"       # red        — untaught / failed
MAP_MISSING_HEX = FILL_WARN_HEX   # amber  — expected here, never attempted
MAP_INFO_HEX = "BDD7EE"       # blue   — real rack, but this report scores nothing here
                              #          (another area, or expected yet never probed)
MAP_IGNORED_HEX = "D9D9D9"    # grey   — nothing to report (unreachable / out of scope)

# Magnitude heat scale (orange low → yellow → green high, solid red over spec).
SCALE_LOW_RGB = (0xFF, 0xC0, 0x00)
SCALE_MID_RGB = (0xFF, 0xEB, 0x84)
SCALE_HIGH_RGB = (0x63, 0xBE, 0x7B)
SCALE_FAIL_HEX = "FF0000"

# Diverging heat scale for signed deviations. Blue/red with a near-white midpoint reads
# correctly under the common red-green colour vision deficiency (unlike red/green).
DIVERGING_LOW = "FF4472C4"    # blue      — below the reference
DIVERGING_MID = "FFF2F2F2"    # near-white — on the reference
DIVERGING_HIGH = "FFE06666"   # red       — above the reference

# Excel's own green-yellow-red 3-colour scale, for maps where low is good and high is bad
# (measured clearances). Same hues as the magnitude scale, reversed.
GYR_LOW = "FF63BE7B"
GYR_MID = "FFFFEB84"
GYR_HIGH = "FFF8696B"

# Column widths.
WIDTH_LABEL_COL = 12          # the "Level \ Column" label column of any matrix
WIDTH_VALUE_COL = 8           # a matrix cell holding a rounded value
WIDTH_MAP_COL = 5.4           # a full-rack (72-column) map cell holding a value
WIDTH_MAP_FILL_COL = 3.2      # a full-rack map cell that is colour only

EM_DASH = "—"            # written in place of "no data" in a matrix cell

MATRIX_CORNER = "Level \\ Column"

_FORMULA_TRIGGER_PREFIX = frozenset("=+-@")

_STYLE_CACHE: dict[str, Any] = {}


# ======================================================================================
# Style objects
# ======================================================================================


def styles() -> dict[str, Any]:
    if not _STYLE_CACHE:
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

        thin = Side(style="thin", color=GRID_GRAY)
        _STYLE_CACHE.update(
            {
                "hdr_font": Font(bold=True, color="FFFFFF"),
                "hdr_fill": PatternFill("solid", fgColor=NAVY),
                "title_font": Font(bold=True, size=14, color=NAVY),
                "section_font": Font(bold=True, size=11, color=NAVY),
                "subtitle_font": Font(italic=True, size=9, color=SUBTITLE_GRAY),
                "spec_font": Font(bold=True, color=NAVY),
                "center": Alignment(horizontal="center"),
                "border": Border(left=thin, right=thin, top=thin, bottom=thin),
                "fill_ok": PatternFill("solid", fgColor=FILL_OK_HEX),
                "fill_warn": PatternFill("solid", fgColor=FILL_WARN_HEX),
                "fill_fail": PatternFill("solid", fgColor=FILL_FAIL_HEX),
                "fill_review": PatternFill("solid", fgColor=FILL_REVIEW_HEX),
                "fill_spec": PatternFill("solid", fgColor=FILL_SPEC_HEX),
                "fill_out_of_tol": PatternFill("solid", fgColor=FILL_OUT_OF_TOL_HEX),
                "map_pass": PatternFill("solid", fgColor=MAP_PASS_HEX),
                "map_fail": PatternFill("solid", fgColor=MAP_FAIL_HEX),
                "map_missing": PatternFill("solid", fgColor=MAP_MISSING_HEX),
                "map_info": PatternFill("solid", fgColor=MAP_INFO_HEX),
                "map_ignored": PatternFill("solid", fgColor=MAP_IGNORED_HEX),
            }
        )
    return _STYLE_CACHE


def cell_value(v: Any) -> Any:
    if isinstance(v, str) and v and v[0] in _FORMULA_TRIGGER_PREFIX:
        return "'" + v
    return v


# ======================================================================================
# Headings, headers, labels
# ======================================================================================


def title(ws: Any, text: str, row: int, col: int = 1, kind: str = "title", span: int = 1) -> Any:
    s = styles()
    c = ws.cell(row=row, column=col, value=cell_value(text))
    c.font = s.get(f"{kind}_font", s["title_font"])
    if span > 1:
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + span - 1)
    return c


def header_row(ws: Any, headers: Sequence[Any], row: int = 1, start_col: int = 1) -> int:
    for j, h in enumerate(headers, start=start_col):
        label_cell(ws, row, j, h)
    return row + 1


def label_cell(ws: Any, row: int, col: int, value: Any) -> Any:
    s = styles()
    c = ws.cell(row=row, column=col, value=cell_value(value))
    c.font = s["hdr_font"]
    c.fill = s["hdr_fill"]
    c.alignment = s["center"]
    c.border = s["border"]
    return c


def set_widths(ws: Any, widths: Sequence[float], start_col: int = 1) -> None:
    from openpyxl.utils import get_column_letter

    for i, w in enumerate(widths, start=start_col):
        ws.column_dimensions[get_column_letter(i)].width = w


def set_matrix_widths(
    ws: Any,
    n_cols: int,
    cell_width: float = WIDTH_VALUE_COL,
    label_width: float = WIDTH_LABEL_COL,
) -> None:
    set_widths(ws, [label_width] + [cell_width] * n_cols)


def finish_sheet(ws: Any, freeze: str | None = None) -> None:
    if freeze and ws.freeze_panes is None:
        ws.freeze_panes = freeze
    ws.sheet_view.showGridLines = False


# ======================================================================================
# Level ordering — descending everywhere
# ======================================================================================


def _level_key(v: Any) -> tuple[int, Any]:
    try:
        return (0, float(v))
    except (TypeError, ValueError):
        return (1, str(v))


def sort_levels(levels) -> list:
    """Descending: level 16 first, level 1 last."""
    return sorted(set(levels), key=_level_key, reverse=True)


def level_label(level: Any) -> str:
    return f"L{level}"


def sort_level_rows(
    items: Sequence[Any],
    level_of: Callable[[Any], Any],
    then_by: Callable[[Any], Any] | None = None,
) -> list:
    ordered = sorted(items, key=then_by) if then_by is not None else list(items)
    # Stable sort, so `then_by` survives as the ascending tiebreak inside each level.
    ordered.sort(key=lambda it: _level_key(level_of(it)), reverse=True)
    return ordered


# ======================================================================================
# Tables
# ======================================================================================


def write_table(
    ws: Any,
    start_row: int,
    headers: Sequence[Any],
    rows: Sequence[Sequence[Any]],
    fill: Any = None,
    start_col: int = 1,
) -> int:
    s = styles()
    r = header_row(ws, headers, row=start_row, start_col=start_col)
    for row in rows:
        for c, v in enumerate(row, start=start_col):
            cell = ws.cell(row=r, column=c, value=cell_value(v))
            cell.border = s["border"]
            if fill is not None:
                cell.fill = fill
        r += 1
    return r


class MatrixCell(NamedTuple):
    value: Any = None
    fill: Any = None
    number_format: str | None = None


def write_matrix(
    ws: Any,
    start_row: int,
    row_keys: Sequence[Any],
    col_keys: Sequence[Any],
    cell_fn: Callable[[Any, Any], MatrixCell | None],
    corner: str = MATRIX_CORNER,
    row_label_fn: Callable[[Any], Any] = level_label,
    col_label_fn: Callable[[Any], Any] = lambda c: c,
    start_col: int = 1,
) -> tuple[int, int]:
    s = styles()
    r = start_row
    label_cell(ws, r, start_col, corner)
    for j, ck in enumerate(col_keys, start=start_col + 1):
        label_cell(ws, r, j, col_label_fn(ck))
    r += 1
    first_data_row = r

    for rk in row_keys:
        label_cell(ws, r, start_col, row_label_fn(rk))
        for j, ck in enumerate(col_keys, start=start_col + 1):
            cell = ws.cell(row=r, column=j)
            cell.border = s["border"]
            mc = cell_fn(rk, ck)
            if mc is None:
                continue
            if mc.value is not None:
                cell.value = cell_value(mc.value)
                cell.alignment = s["center"]
            if mc.fill is not None:
                cell.fill = mc.fill
            if mc.number_format is not None:
                cell.number_format = mc.number_format
        r += 1
    return r, first_data_row


# ======================================================================================
# Heat scales
# ======================================================================================


def scale_bounds(values: Sequence[float]) -> tuple[float, float]:
    """(0.0, 0.0) if `values` is empty."""
    if not values:
        return 0.0, 0.0
    return min(values), max(values)


def scale_gradient_hex(t: float) -> str:
    """Magnitude gradient: t=0 → orange (low), t=0.5 → yellow, t=1 → green (high)."""
    t = max(0.0, min(1.0, t))
    if t <= 0.5:
        lo, hi, f = SCALE_LOW_RGB, SCALE_MID_RGB, t / 0.5
    else:
        lo, hi, f = SCALE_MID_RGB, SCALE_HIGH_RGB, (t - 0.5) / 0.5
    r, g, b = (round(a + (c - a) * f) for a, c in zip(lo, hi))
    return f"{r:02X}{g:02X}{b:02X}"


def scale_fill(value: float | None, vmin: float, vmax: float, failed: bool) -> Any:
    """Magnitude fill: failed → solid red; else orange(low)→green(high) gradient by value."""
    from openpyxl.styles import PatternFill

    if value is None:
        return None
    if failed:
        return PatternFill("solid", fgColor=SCALE_FAIL_HEX)
    t = 0.0 if vmax <= vmin else (value - vmin) / (vmax - vmin)
    return PatternFill("solid", fgColor=scale_gradient_hex(t))


def symmetric_bound(deviations: Sequence[float], floor: float = 1.0) -> float:
    """Symmetric colour-scale half-width: 98th percentile |dev|, so outliers don't wash it out."""
    if not deviations:
        return floor
    d = sorted(abs(x) for x in deviations)
    p98 = d[min(len(d) - 1, int(len(d) * 0.98))]
    return max(floor, round(p98 + 0.05, 1))


def add_color_scale(
    ws: Any,
    first_row: int,
    last_row: int,
    first_col: int,
    last_col: int,
    low: float,
    mid: float,
    high: float,
    colors: tuple[str, str, str] = (DIVERGING_LOW, DIVERGING_MID, DIVERGING_HIGH),
) -> None:
    """Endpoints are fixed numbers, never min/max, so every block on a sheet shares one scale
    and stays comparable to the others. Text cells (an 'X' marker, an em dash) are skipped by
    Excel, so a marker can never read as a scaled value.
    """
    from openpyxl.formatting.rule import ColorScaleRule
    from openpyxl.utils import get_column_letter

    if last_row < first_row or last_col < first_col:
        return
    rng = (
        f"{get_column_letter(first_col)}{first_row}:"
        f"{get_column_letter(last_col)}{last_row}"
    )
    lo_c, mid_c, hi_c = colors
    ws.conditional_formatting.add(
        rng,
        ColorScaleRule(
            start_type="num", start_value=low, start_color=lo_c,
            mid_type="num", mid_value=mid, mid_color=mid_c,
            end_type="num", end_value=high, end_color=hi_c,
        ),
    )
