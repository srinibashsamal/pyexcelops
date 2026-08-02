"""
pyexcelops.py
==================
A utility for writing multiple pandas DataFrames to a single, consistently
formatted Excel workbook (.xlsx).

Public API
----------
    save_formatted_excel(...)   - the only function you need

Typical usage
-------------
    from excel_ops import save_formatted_excel

    save_formatted_excel(
        dfs=[df_sales, df_returns],
        output_path="report.xlsx",
        sheet_names=["Sales", "Returns"],
        highlight_cols_list=[["Revenue"], []],
        as_table=True,
        auto_fit_row_heights=True,
    )
"""

from __future__ import annotations

import re
import math
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_LIGHT_BLUE: str = "ADCFEB"
_NAVY_BLUE: str = "002266"
_FONT_NAME: str = "Calibri"
_DEFAULT_TABLE_STYLE: str = "TableStyleMedium1"
_NO_DATA_MESSAGE: str = "This sheet contains no data based on the current inputs!"
_FILTER_ICON_WIDTH: float = 3


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _make_styles() -> Dict[str, Any]:
    """Return a dict of reusable openpyxl style objects."""
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    return {
        "header_font": Font(name=_FONT_NAME, bold=True, color=_NAVY_BLUE),
        "header_fill": PatternFill(
            start_color=_LIGHT_BLUE, end_color=_LIGHT_BLUE, fill_type="solid"
        ),
        "header_align": Alignment(
            horizontal="center", vertical="center", wrap_text=True
        ),
        "body_wrap_align": Alignment(vertical="top", wrap_text=True),
        "thin_border": border,
    }


def _calc_column_width(
    series: pd.Series,
    header: str,
    padding: int = 2,
    max_width: int = 50,
) -> int:
    """Return a best-fit column width (in Excel character units), capped at *max_width*."""
    if series.empty:
        return min(len(header) + padding, max_width)

    max_len = max(series.astype(str).str.len().max(), len(header))
    return min(max_len + padding, max_width)


def _sanitize_table_name(name: str) -> str:
    """
    Produce a valid Excel table ``displayName``:
    no spaces, must start with a letter or underscore, max 255 chars.
    """
    base = re.sub(r"[^A-Za-z0-9_]", "_", name.strip()) or "Table"
    if not re.match(r"^[A-Za-z_]", base):
        base = "_" + base
    return base[:255]


def _unique_table_name(workbook, desired: str) -> str:
    """Return *desired* if it is not already used in *workbook*, otherwise append ``_2``, ``_3``, …"""
    existing: set[str] = set()
    for ws in workbook.worksheets:
        existing.update(ws.tables.keys())
    if desired not in existing:
        return desired
    i = 2
    while f"{desired}_{i}" in existing:
        i += 1
    return f"{desired}_{i}"


def _apply_wrap_row_heights(
    ws,
    min_row: int,
    max_row: int,
    wrap_cols: List[int],
    base_line_height: float = 15.0,
    min_height: float = 15.0,
    max_height: float = 240.0,
) -> None:
    """
    Heuristically set row heights for rows that contain wrapped text.

    Estimates the number of visible lines per cell from the column width and
    the length of the cell value, then sets the row height accordingly.
    """
    if not wrap_cols:
        return

    # Cache column widths (Excel default ≈ 8.43 if not explicitly set).
    col_widths: Dict[int, float] = {}
    for col_idx in wrap_cols:
        raw = ws.column_dimensions[get_column_letter(col_idx)].width
        col_widths[col_idx] = float(raw) if raw else 8.43

    fudge = 1.05  # slight under-estimation guard

    for row_num in range(min_row, max_row + 1):
        needed_lines = 1

        for col_idx in wrap_cols:
            cell = ws.cell(row=row_num, column=col_idx)
            text = str(cell.value) if cell.value is not None else ""
            if not text:
                continue

            chars_per_line = max(int(col_widths[col_idx] / fudge), 5)
            lines_for_cell = sum(
                max(1, math.ceil(len(para) / chars_per_line)) if para else 1
                for para in text.split("\n")
            )
            needed_lines = max(needed_lines, lines_for_cell)

        ws.row_dimensions[row_num].height = max(
            min_height, min(max_height, needed_lines * base_line_height)
        )


def _write_empty_sheet(
    ws, ncols: int, styles: Dict[str, Any], wrap_threshold: int
) -> None:
    """Write a 'no data' notice on a sheet that has headers but zero data rows."""
    if ncols == 0:
        ws["A1"] = _NO_DATA_MESSAGE
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws["A1"].font = Font(name=_FONT_NAME, italic=True, color=_NAVY_BLUE)
        return

    # Write message in row 2, merged across all columns.
    msg_cell = ws.cell(row=2, column=1)
    msg_cell.value = _NO_DATA_MESSAGE
    msg_cell.alignment = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )
    msg_cell.font = Font(name=_FONT_NAME, italic=True, color=_NAVY_BLUE)
    ws.merge_cells(f"A2:{get_column_letter(ncols)}2")
    ws.row_dimensions[2].height = 20

    for row in ws.iter_rows(min_row=1, max_row=2, max_col=ncols):
        for cell in row:
            cell.border = styles["thin_border"]

    for col_idx, header_cell in enumerate(ws[1], start=1):
        width = min(len(str(header_cell.value)) + 2, wrap_threshold * 2)
        ws.column_dimensions[get_column_letter(col_idx)].width = width


# ---------------------------------------------------------------------------
# Validators
# ---------------------------------------------------------------------------


def _ensure_sheet_names(
    dfs: List[pd.DataFrame], sheet_names: Optional[List[str]]
) -> List[str]:
    """
    Ensure sheet_names list matches the number of DataFrames.
    Fills missing names with 'Sheet1', 'Sheet2', etc.
    """
    final_names = list(sheet_names or [])
    num_dfs = len(dfs)

    for i in range(len(final_names), num_dfs):
        final_names.append(f"Sheet{i+1}")

    return final_names[:num_dfs]


def _validate_highlight_cols(
    dfs: List[pd.DataFrame],
    highlight_cols_list: Optional[List[Optional[List[str]]]],
) -> List[Optional[List[str]]]:
    """Ensure highlight_cols_list matches dfs length."""

    if highlight_cols_list is None:
        return [None] * len(dfs)

    if len(highlight_cols_list) != len(dfs):
        raise ValueError(
            f"'highlight_cols_list' length ({len(highlight_cols_list)}) "
            f"must match 'dfs' length ({len(dfs)})."
        )

    return highlight_cols_list


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def save_formatted_excel(
    dfs: List[pd.DataFrame],
    output_path: Union[str, Path],
    sheet_names: Optional[List[str]] = None,
    highlight_cols_list: Optional[List[Optional[List[str]]]] = None,
    freeze_header: bool = False,
    wrap_threshold: int = 40,
    as_table: bool = False,
    table_style: str = _DEFAULT_TABLE_STYLE,
    auto_fit_row_heights: bool = False,
    show_gridlines: bool = True,
    add_filters: bool = False,
) -> None:
    """Write multiple DataFrames to a single, consistently formatted Excel workbook.

    Each DataFrame is placed on its own worksheet with standardized header
    styling (bold navy text on light-blue fill), thin cell borders, and
    best-effort auto-fitted column widths.  A range of optional enhancements
    can be enabled via keyword arguments.

    Args:
        dfs (List[pd.DataFrame]): DataFrames to write; one per sheet.
        output_path (Union[str, Path]): Destination ``.xlsx`` file path
            (e.g. ``"report.xlsx"`` or a ``Path``).
        sheet_names (Optional[List[str]]): Worksheet names. If `None` or shorter than
            dfs, defaults like "Sheet1", "Sheet2" are used.
        highlight_cols_list (Optional[List[Optional[List[str]]]], optional):
            Per-sheet list of column names whose data cells should receive the same
            light-blue fill as the header.

            Pass ``None`` or an empty list for sheets that need no highlighting.
            Defaults to `None` no highlighting.
        freeze_header (bool, optional): Freeze the first row so headers stay
            visible while scrolling. Defaults to `False`.
        wrap_threshold (int, optional): Column-width threshold (Excel character units)
            above which body cells receive ``wrap_text=True``. Defaults to `40`.
        as_table (bool, optional): Convert each sheet's data range to a native
            Excel Table (banded rows, built-in filter dropdowns). Defaults to `False`.
        table_style (str, optional): Excel Table style name, e.g. ``"TableStyleMedium2"``.
            Only used when *as_table* is ``True``. Defaults to `_DEFAULT_TABLE_STYLE`.
        auto_fit_row_heights (bool, optional): Apply heuristic row-height adjustment
            for rows containing wrapped text. Defaults to `False`.
        show_gridlines (bool, optional): Show (``True``) or hide (``False``)
            sheet gridlines. Defaults to `True`.
        add_filters (bool, optional): Add AutoFilter dropdowns to the header row when
            *as_table* is ``False``.

            Ignored when *as_table* is ``True``
            (tables include filters automatically). Defaults to `False`.

    Raises:
        ValueError: If *highlight_cols_list* has a different length than *dfs*.

    Examples
    --------
    >>> save_formatted_excel(
    ...     dfs=[df_due, df_overdue],
    ...     output_path="My_Report.xlsx",
    ...     sheet_names=["Due", "Overdue"],
    ...     highlight_cols_list=[["Protocol Name"], []],
    ...     freeze_header=True,
    ...     wrap_threshold=25,
    ...     as_table=True,
    ...     table_style="TableStyleMedium2",
    ...     auto_fit_row_heights=True,
    ...     show_gridlines=False,
    ... )

    """

    output_path = Path(output_path)

    # Check if the directory exists
    if not output_path.parent.exists():
        raise FileNotFoundError(f"The directory '{output_path.parent}' does not exist.")

    # Validation and Normalization
    sheet_names = _ensure_sheet_names(dfs, sheet_names)
    highlight_cols_list = _validate_highlight_cols(dfs, highlight_cols_list)

    styles = _make_styles()

    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for df, sheet_name, highlight_cols in zip(
                dfs, sheet_names, highlight_cols_list
            ):
                highlight_set = set(highlight_cols or [])
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]

                ws.sheet_view.showGridLines = bool(show_gridlines)
                if freeze_header:
                    ws.freeze_panes = "A2"

                nrows, ncols = len(df), len(df.columns)

                # ---- Empty DataFrame ----
                if nrows == 0:
                    _write_empty_sheet(ws, ncols, styles, wrap_threshold)
                    continue

                # ---- Header styling & column widths ----
                wrapped_col_indices: List[int] = []

                filter_active = as_table or add_filters
                filter_padding = _FILTER_ICON_WIDTH if filter_active else 0.0

                for col_idx, header_cell in enumerate(ws[1], start=1):
                    header_cell.font = styles["header_font"]
                    header_cell.fill = styles["header_fill"]
                    header_cell.alignment = styles["header_align"]
                    header_cell.border = styles["thin_border"]

                    # series = df.iloc[:, col_idx - 1] if not df.empty else pd.Series()
                    series = df.iloc[:, col_idx - 1]
                    width = _calc_column_width(
                        series, str(header_cell.value), max_width=wrap_threshold * 2
                    )
                    ws.column_dimensions[get_column_letter(col_idx)].width = (
                        width + filter_padding
                    )

                    if width >= wrap_threshold:
                        wrapped_col_indices.append(col_idx)
                        for row in ws.iter_rows(
                            min_row=2,
                            min_col=col_idx,
                            max_col=col_idx,
                            max_row=nrows + 1,
                        ):
                            row[0].alignment = styles["body_wrap_align"]

                # ---- Borders (full data range) ----
                for row in ws.iter_rows(min_row=1, max_row=nrows + 1, max_col=ncols):
                    for cell in row:
                        cell.border = styles["thin_border"]

                # ---- Column highlighting ----
                headers = [cell.value for cell in ws[1]]
                for col_idx, name in enumerate(headers, start=1):
                    if name in highlight_set:
                        for row_num in range(2, nrows + 2):
                            ws.cell(row=row_num, column=col_idx).fill = styles[
                                "header_fill"
                            ]

                # ---- Heuristic row heights ----
                if auto_fit_row_heights and wrapped_col_indices:
                    _apply_wrap_row_heights(
                        ws,
                        min_row=2,
                        max_row=nrows + 1,
                        wrap_cols=wrapped_col_indices,
                    )

                # ---- Excel Table or plain AutoFilter ----
                if as_table:
                    last_col = get_column_letter(ncols)
                    last_row = nrows + 1
                    ref = f"A1:{last_col}{last_row}"

                    desired = _sanitize_table_name(f"{sheet_name}_Table")
                    tname = _unique_table_name(writer.book, desired)

                    tab = Table(displayName=tname, ref=ref)
                    tab.tableStyleInfo = TableStyleInfo(
                        name=table_style,
                        showFirstColumn=False,
                        showLastColumn=False,
                        showRowStripes=True,
                        showColumnStripes=False,
                    )
                    ws.add_table(tab)

                elif add_filters:
                    ws.auto_filter.ref = f"A1:{get_column_letter(ncols)}{nrows + 1}"

        print(f"\nFile saved and formatted: {output_path} ")
        for name, df in zip(sheet_names, dfs):
            print(f"   └─ Sheet '{name}': {df.shape[0]} rows x {df.shape[1]} cols")

    except Exception as exc:
        print(f"Failed to save '{output_path}': {exc}")
        raise
