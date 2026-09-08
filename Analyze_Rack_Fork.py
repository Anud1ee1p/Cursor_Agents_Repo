from __future__ import annotations
import json
import os
import re
import statistics
import sys
from collections import defaultdict
from typing import Any, Literal

from report_schema import (
    EM_DASH,
    FILL_FAIL_HEX,
    FILL_OK_HEX,
    FILL_WARN_HEX,
    NAVY,
    SUBTITLE_GRAY,
    MatrixCell,
    cell_value as _cell_value_for_excel,
    finish_sheet as _finish_sheet,
    header_row as _style_header,
    level_label as _level_label,
    scale_bounds as _scale_bounds,
    scale_fill as _scale_fill,
    set_matrix_widths as _set_matrix_widths,
    set_widths as _set_widths,
    sort_level_rows as _sort_level_rows,
    sort_levels as _sort_levels,
    styles as _styles,
    title as _title,
    write_matrix as _write_matrix,
    write_table as _write_styled_table,
)

# ======================================================================================
# Constants
# ======================================================================================

LocationSelection = Literal["ALL"] | frozenset[str]

TRAVEL_F2_AHEAD_OF_F1_MM = 1180.0
TRAVEL_TARGET = TRAVEL_F2_AHEAD_OF_F1_MM


def _output_basename(aisle: int | None) -> str:
    if aisle is not None:
        return f"Aisle{aisle}_RackData_Analysis_Report.xlsx"
    return "RackData_Analysis_Report.xlsx"

NON_CD_CODE = "10"
# Fork-unreachable columns (2-digit key codes): Fork 1 cannot reach 01/02, Fork 2 cannot
# reach 71/72. The JSON stores zero-filled placeholder fork positions there, so any
# Fork2 − Fork1 delta at these columns is meaningless on EVERY side (CD and Non-CD alike).
FORK_UNREACHABLE_COL_CODES = frozenset({"01", "02", "71", "72"})
CD_SIDE_ALT_SUFFIX = "30"
CD_SIDE_DEFAULT_SUFFIX = "40"
CD_SIDE_ALT_AISLES = frozenset({"02", "05"})

LOCATION_NAMES: dict[str, str] = {
    "01": "Aisle Buffer",
    "02": "CD",
    "03": "RTA",
    "04": "HTA",
    "05": "OCV",
}

REQUIRED_KEYS = (
    "Fork1TravelPos",
    "Fork1HoistPos",
    "Fork1Pos",
    "Fork2TravelPos",
    "Fork2HoistPos",
    "Fork2Pos",
)

LASER_PADDLE_OFFSET_KEYS = (
    "LaserToPaddle1X",
    "LaserToPaddle2X",
    "LaserToPaddle1Y",
    "LaserToPaddle2Y",
    "LaserToPaddle1Z",
    "LaserToPaddle2Z",
)

PADDLE_OFFSET_LOCATIONS = frozenset({"01", "03", "04"})

# Offset entry order (one value per row): group, field label, storage key, side index
# (index 0 = Non-CD / Aging sides 10/20, index 1 = CD sides 30/40).
OFFSET_ENTRY_SEQUENCE: tuple[tuple[str, str, str, int], ...] = (
    ("Fork 1 — CD side", "Travel (X)", "LaserToPaddle1X", 1),
    ("Fork 1 — CD side", "Hoist (Z)", "LaserToPaddle1Z", 1),
    ("Fork 1 — CD side", "Insert (Y)", "LaserToPaddle1Y", 1),
    ("Fork 1 — Non-CD side", "Travel (X)", "LaserToPaddle1X", 0),
    ("Fork 1 — Non-CD side", "Hoist (Z)", "LaserToPaddle1Z", 0),
    ("Fork 1 — Non-CD side", "Insert (Y)", "LaserToPaddle1Y", 0),
    ("Fork 2 — CD side", "Travel (X)", "LaserToPaddle2X", 1),
    ("Fork 2 — CD side", "Hoist (Z)", "LaserToPaddle2Z", 1),
    ("Fork 2 — CD side", "Insert (Y)", "LaserToPaddle2Y", 1),
    ("Fork 2 — Non-CD side", "Travel (X)", "LaserToPaddle2X", 0),
    ("Fork 2 — Non-CD side", "Hoist (Z)", "LaserToPaddle2Z", 0),
    ("Fork 2 — Non-CD side", "Insert (Y)", "LaserToPaddle2Y", 0),
)

OFFSET_FORMAT_EXAMPLE = (
    "Enter one value per field in mm (e.g. Travel ~140, Hoist ~ -140 to -260, Insert ~0). "
    "Values are saved per aisle."
)

OFFSETS_SETTINGS_FILENAME = ".analyze_rack_fork_offsets.json"
AISLE_NUMBERS = tuple(range(1, 10))  # aisles 1–9

LaserPaddleOffsets = dict[str, list[float]]


# ======================================================================================
# Fixture heat matrices
# ======================================================================================


def _write_heat_matrix_sections(
    ws: Any, start_row: int, section_title: str, cells: list[dict], tol: float
) -> int:
    r = start_row
    _title(ws, section_title, r, kind="title", span=6)
    r += 2

    area_name = {3: "RTA", 4: "HTA"}
    groups: dict[tuple[int, int, int], list[dict]] = defaultdict(list)
    for c in cells:
        groups[(c["loc"], c["line"], c["fork"])].append(c)

    freeze_cell: str | None = None
    for loc in (3, 4):
        for line in (10, 40):
            for fork in (0, 1, 2):
                grp = groups.get((loc, line, fork))
                if not grp:
                    continue
                col_labels = [lbl for _, lbl in sorted({(g["csort"], g["clabel"]) for g in grp})]
                row_labels = [
                    lbl for _, lbl in sorted({(g["rsort"], g["rlabel"]) for g in grp}, reverse=True)
                ]
                grid = {(g["rlabel"], g["clabel"]): g["error"] for g in grp}

                errs = [g["error"] for g in grp]
                failed = sum(1 for e in errs if e > tol)
                n = len(errs)
                vmin, vmax = _scale_bounds([e for e in errs if e <= tol])
                side = "Aging" if line == 10 else "CD"
                label = f"{area_name[loc]} · {side}" + (f" · Fork {fork}" if fork else "")
                strip = (
                    f"{label}    spec ≤ {tol:g} mm    n {n} · failed {failed} · "
                    f"pass {100.0 * (n - failed) / n:.1f}%    mean {statistics.fmean(errs):.2f}  "
                    f"max {max(errs):.2f}"
                )
                _title(ws, strip, r, kind="section", span=1 + len(col_labels))
                r += 1

                def cell_fn(rlabel: Any, clabel: Any, _grid=grid, _lo=vmin, _hi=vmax) -> MatrixCell:
                    err = _grid.get((rlabel, clabel))
                    if err is None:
                        return MatrixCell(value=EM_DASH)
                    return MatrixCell(round(err, 2), _scale_fill(err, _lo, _hi, err > tol))

                r, first_data_row = _write_matrix(
                    ws, r, row_labels, col_labels, cell_fn,
                    corner="Level \\ Col", row_label_fn=lambda lbl: lbl,
                )
                if freeze_cell is None:
                    freeze_cell = f"B{first_data_row}"
                r += 1

    if not any(groups.values()):
        _title(ws, "No data for this check in the current filter.", r, kind="subtitle", span=6)
        r += 1
    _finish_sheet(ws, freeze_cell)
    return r + 1


# ======================================================================================
# Step 2 — Laser-to-paddle offsets
# ======================================================================================


def _settings_path() -> str:
    try:
        base = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        base = os.getcwd()
    return os.path.join(base, OFFSETS_SETTINGS_FILENAME)


def _validate_offsets(raw: Any) -> LaserPaddleOffsets | None:
    if not isinstance(raw, dict):
        return None
    out: LaserPaddleOffsets = {}
    for key in LASER_PADDLE_OFFSET_KEYS:
        pair = raw.get(key)
        if not isinstance(pair, list) or len(pair) != 2:
            return None
        try:
            out[key] = [float(pair[0]), float(pair[1])]
        except (TypeError, ValueError):
            return None
    return out


def _load_all_offsets() -> dict[str, LaserPaddleOffsets]:
    try:
        with open(_settings_path(), encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(data, dict):
        return {}
    out: dict[str, LaserPaddleOffsets] = {}
    for aisle_key, raw in data.items():
        validated = _validate_offsets(raw)
        if validated is not None:
            out[str(aisle_key)] = validated
    return out


def _load_offsets_for_aisle(aisle: int) -> LaserPaddleOffsets | None:
    return _load_all_offsets().get(str(aisle))


def _save_offsets_for_aisle(aisle: int, offsets: LaserPaddleOffsets) -> None:
    all_offsets = _load_all_offsets()
    all_offsets[str(aisle)] = offsets
    payload = {
        a: {key: vals[key] for key in LASER_PADDLE_OFFSET_KEYS}
        for a, vals in sorted(all_offsets.items(), key=lambda kv: int(kv[0]) if kv[0].isdigit() else 0)
    }
    try:
        with open(_settings_path(), "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2)
    except OSError:
        pass


def _offsets_to_entry_values(offsets: LaserPaddleOffsets | None) -> dict[tuple[str, int], str]:
    if offsets is None:
        return {}
    return {(key, idx): f"{offsets[key][idx]:g}" for _grp, _field, key, idx in OFFSET_ENTRY_SEQUENCE}


def _offsets_summary_text(offsets: LaserPaddleOffsets) -> str:
    lines: list[str] = []
    last_group = None
    for group, field, key, idx in OFFSET_ENTRY_SEQUENCE:
        if group != last_group:
            lines.append(f"{group}:")
            last_group = group
        lines.append(f"    {field}: {offsets[key][idx]:g}")
    return "\n".join(lines)


def _offsets_from_entry_values(values: dict[tuple[str, int], str]) -> LaserPaddleOffsets:
    out: LaserPaddleOffsets = {key: [0.0, 0.0] for key in LASER_PADDLE_OFFSET_KEYS}
    for group, field, key, idx in OFFSET_ENTRY_SEQUENCE:
        raw = values.get((key, idx), "").strip()
        if not raw:
            raise ValueError(f"Enter a value for {group} — {field}.")
        try:
            out[key][idx] = float(raw)
        except ValueError:
            raise ValueError(f"{group} — {field}: must be a number.")
    return out


def _center_tk_window(root: Any, width: int | None = None, height: int | None = None) -> None:
    root.update_idletasks()
    w = width if width is not None else root.winfo_reqwidth()
    h = height if height is not None else root.winfo_reqheight()
    x = max(0, (root.winfo_screenwidth() - w) // 2)
    y = max(0, (root.winfo_screenheight() - h) // 2)
    root.geometry(f"{w}x{h}+{x}+{y}")


def _edit_offsets_gui(aisle: int, prefill: LaserPaddleOffsets | None) -> LaserPaddleOffsets | None:
    import tkinter as tk
    from tkinter import messagebox, ttk

    root = tk.Tk()
    root.title(f"Rack analysis — offsets for Aisle {aisle}")
    root.attributes("-topmost", True)
    root.resizable(False, False)

    main = ttk.Frame(root, padding=14)
    main.pack(fill="both", expand=True)

    ttk.Label(
        main,
        text=f"Laser-to-paddle offsets — Aisle {aisle}",
        font=("Segoe UI", 12, "bold"),
    ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))

    ttk.Label(
        main,
        text=(
            "Enter/adjust this aisle's calibration (one value per field, mm).\n"
            "Applied to locations 01 (Aisle Buffer), 03 (RTA), and 04 (HTA); saved per aisle."
        ),
        justify="left",
        wraplength=360,
    ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 10))

    prefill_values = _offsets_to_entry_values(prefill)
    entry_widgets: dict[tuple[str, int], ttk.Entry] = {}
    grid_row = 2
    last_group = None
    for group, field, key, idx in OFFSET_ENTRY_SEQUENCE:
        if group != last_group:
            ttk.Label(main, text=group, font=("Segoe UI", 9, "bold")).grid(
                row=grid_row, column=0, columnspan=2, sticky="w", pady=(8, 2)
            )
            grid_row += 1
            last_group = group
        ttk.Label(main, text=field).grid(row=grid_row, column=0, sticky="w", padx=(12, 12), pady=3)
        entry = ttk.Entry(main, width=16, justify="right")
        entry.grid(row=grid_row, column=1, padx=4, pady=3)
        value = prefill_values.get((key, idx), "")
        if value:
            entry.insert(0, value)
        entry_widgets[(key, idx)] = entry
        grid_row += 1

    ttk.Label(
        main, text=OFFSET_FORMAT_EXAMPLE, foreground="#555555", wraplength=360, font=("Segoe UI", 8)
    ).grid(row=grid_row, column=0, columnspan=2, sticky="w", pady=(10, 12))
    grid_row += 1

    chosen: list[LaserPaddleOffsets | None] = [None]

    def _snapshot_form() -> dict[tuple[str, int], str]:
        return {ident: entry.get() for ident, entry in entry_widgets.items()}

    def apply_offsets() -> None:
        try:
            offsets = _offsets_from_entry_values(_snapshot_form())
        except ValueError as e:
            messagebox.showerror("Invalid offsets", str(e), parent=root)
            return
        _save_offsets_for_aisle(aisle, offsets)
        chosen[0] = offsets
        root.destroy()

    def skip_offsets() -> None:
        chosen[0] = None
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", skip_offsets)

    btn_row = ttk.Frame(main)
    btn_row.grid(row=grid_row, column=0, columnspan=2, sticky="e")
    ttk.Button(btn_row, text="No offsets", command=skip_offsets, width=14).pack(side=tk.LEFT, padx=(0, 8))
    ttk.Button(btn_row, text="Apply offsets", command=apply_offsets, width=14).pack(side=tk.LEFT)

    root.bind("<Return>", lambda _event: apply_offsets())
    root.bind("<Escape>", lambda _event: skip_offsets())
    first = OFFSET_ENTRY_SEQUENCE[0]
    entry_widgets[(first[2], first[3])].focus_set()
    _center_tk_window(root)
    root.mainloop()
    return chosen[0]


def _confirm_saved_offsets_gui(aisle: int, saved: LaserPaddleOffsets) -> str:
    import tkinter as tk
    from tkinter import ttk

    root = tk.Tk()
    root.title(f"Rack analysis — Aisle {aisle} offsets")
    root.attributes("-topmost", True)
    root.resizable(False, False)

    main = ttk.Frame(root, padding=14)
    main.pack(fill="both", expand=True)
    ttk.Label(
        main, text=f"Aisle {aisle} has saved offsets:", font=("Segoe UI", 11, "bold")
    ).pack(anchor="w", pady=(0, 6))
    ttk.Label(main, text=_offsets_summary_text(saved), justify="left", font=("Consolas", 9)).pack(
        anchor="w", pady=(0, 6)
    )
    ttk.Label(
        main, text="Apply these, edit them, or run with no offsets?", justify="left", wraplength=420
    ).pack(anchor="w", pady=(0, 10))

    result: list[str] = ["none"]

    def _choose(value: str) -> None:
        result[0] = value
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", lambda: _choose("none"))
    btn_row = ttk.Frame(main)
    btn_row.pack(anchor="e")
    ttk.Button(btn_row, text="No offsets", command=lambda: _choose("none"), width=12).pack(side=tk.LEFT, padx=(0, 8))
    ttk.Button(btn_row, text="Edit", command=lambda: _choose("edit"), width=12).pack(side=tk.LEFT, padx=(0, 8))
    ttk.Button(btn_row, text="Apply", command=lambda: _choose("apply"), width=12).pack(side=tk.LEFT)

    root.bind("<Return>", lambda _event: _choose("apply"))
    root.bind("<Escape>", lambda _event: _choose("none"))
    _center_tk_window(root)
    root.mainloop()
    return result[0]


def _pick_offsets_for_aisle_gui(aisle: int) -> LaserPaddleOffsets | None:
    saved = _load_offsets_for_aisle(aisle)
    if saved is None:
        return _edit_offsets_gui(aisle, None)
    choice = _confirm_saved_offsets_gui(aisle, saved)
    if choice == "apply":
        return saved
    if choice == "edit":
        return _edit_offsets_gui(aisle, saved)
    return None


def _pick_offsets_for_aisle_console(aisle: int) -> LaserPaddleOffsets | None:
    saved = _load_offsets_for_aisle(aisle)
    if saved is not None:
        print(f"Aisle {aisle} has saved offsets:")
        print(_offsets_summary_text(saved))
        while True:
            ans = input("  [A]pply / [E]dit / [N]o offsets: ").strip().lower()
            if ans in ("a", "apply", ""):
                return saved
            if ans in ("n", "no", "none"):
                return None
            if ans in ("e", "edit"):
                break
    else:
        while True:
            ans = input(f"Enter offsets for Aisle {aisle}? (y/n): ").strip().lower()
            if ans in ("n", "no"):
                return None
            if ans in ("y", "yes"):
                break

    print(OFFSET_FORMAT_EXAMPLE)
    prefill = _offsets_to_entry_values(saved)
    values: dict[tuple[str, int], str] = {}
    last_group = None
    for group, field, key, idx in OFFSET_ENTRY_SEQUENCE:
        if group != last_group:
            print(group)
            last_group = group
        default = prefill.get((key, idx), "")
        hint = f" [{default}]" if default.strip() else ""
        entered = input(f"  {field}{hint}: ").strip() or default.strip()
        values[(key, idx)] = entered
    try:
        offsets = _offsets_from_entry_values(values)
        _save_offsets_for_aisle(aisle, offsets)
        return offsets
    except ValueError as e:
        print("ERROR:", e)
        return None


def _paddle_offset_index(side: str) -> int | None:
    if side in ("10", "20"):
        return 0
    if side in ("30", "40"):
        return 1
    return None


def _apply_laser_paddle_offsets(
    location: str,
    side: str,
    f1t: float,
    f2t: float,
    f1h: float,
    f2h: float,
    f1p: float,
    f2p: float,
    offsets: LaserPaddleOffsets,
) -> tuple[float, float, float, float, float, float]:
    if location not in PADDLE_OFFSET_LOCATIONS:
        return f1t, f2t, f1h, f2h, f1p, f2p
    idx = _paddle_offset_index(side)
    if idx is None:
        return f1t, f2t, f1h, f2h, f1p, f2p
    return (
        f1t + offsets["LaserToPaddle1X"][idx],
        f2t + offsets["LaserToPaddle2X"][idx],
        f1h + offsets["LaserToPaddle1Z"][idx],
        f2h + offsets["LaserToPaddle2Z"][idx],
        f1p + offsets["LaserToPaddle1Y"][idx],
        f2p + offsets["LaserToPaddle2Y"][idx],
    )


def _pick_offsets_for_aisle(aisle: int) -> LaserPaddleOffsets | None:
    try:
        return _pick_offsets_for_aisle_gui(aisle)
    except Exception:
        return _pick_offsets_for_aisle_console(aisle)


def _print_offsets_applied(offsets: LaserPaddleOffsets | None) -> None:
    """Echo the exact offsets used this run so cached prefill values are never applied silently."""
    print()
    if offsets is None:
        print("OFFSETS APPLIED THIS RUN: none (raw JSON Travel / Hoist / Pos).")
        return
    print("OFFSETS APPLIED THIS RUN (confirmed in the dialog; [Aging, CD] in mm):")
    for key in LASER_PADDLE_OFFSET_KEYS:
        a, c = offsets[key]
        print(f"  {key:>16}: [{a}, {c}]")
    print(
        "  Applied to Travel (X), Pos (Y), and Hoist (Z) on locations 01/03/04 only "
        "(Aging = sides 10/20, CD = sides 30/40). The All_Rack_Data tab stays raw JSON."
    )


def _expected_cd_side_suffix(location: str) -> str:
    return CD_SIDE_ALT_SUFFIX if location in CD_SIDE_ALT_AISLES else CD_SIDE_DEFAULT_SUFFIX


def _row_is_cd(r: dict) -> bool:
    return r["side"] == _expected_cd_side_suffix(r["location"])


def _row_is_noncd(r: dict) -> bool:
    return r["side"] == NON_CD_CODE


def _is_truthy_inhibit(v: Any) -> bool:
    if isinstance(v, str):
        return v.strip().lower() in ("true", "1", "yes", "y", "t")
    return bool(v)


def _row_is_inhibited(r: dict) -> bool:
    return _is_truthy_inhibit(r.get("inhibit", ""))


# ======================================================================================
# Input selection and helpers
# ======================================================================================


def _pick_rack_json_path() -> str:
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        path = filedialog.askopenfilename(
            title="Select rack data JSON",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        root.destroy()
        if path:
            return path
    except Exception:
        pass

    while True:
        typed = input("Enter the full path to your rack data JSON file: ").strip().strip('"')
        if typed:
            return typed


def _parse_aisle_from_path(path: str) -> int | None:
    m = re.search(r"aisle\s*([0-9]+)", os.path.abspath(path), re.IGNORECASE)
    return int(m.group(1)) if m else None


def _pick_aisle(default: int | None = None) -> int | None:
    detected = f" (detected from folder: {default})" if default else ""
    try:
        import tkinter as tk
        from tkinter import ttk

        root = tk.Tk()
        root.title("Rack analysis — select aisle")
        root.attributes("-topmost", True)
        root.resizable(False, False)
        main = ttk.Frame(root, padding=14)
        main.pack(fill="both", expand=True)
        ttk.Label(
            main,
            text=(
                f"Which aisle is this rack data for?{detected}\n"
                "Even/odd parity selects the CD-side (line 40) rack layout."
            ),
            justify="left",
            wraplength=360,
        ).pack(anchor="w", pady=(0, 10))

        values = [str(a) for a in AISLE_NUMBERS]
        var = tk.StringVar(value=str(default) if default in AISLE_NUMBERS else "")
        combo = ttk.Combobox(main, values=values, textvariable=var, state="readonly", width=10)
        combo.pack(anchor="w", pady=(0, 12))

        chosen: list[int | None] = [None]

        def on_ok() -> None:
            try:
                chosen[0] = int(var.get())
            except ValueError:
                chosen[0] = None
            root.destroy()

        def on_cancel() -> None:
            chosen[0] = None
            root.destroy()

        root.protocol("WM_DELETE_WINDOW", on_cancel)
        btn_row = ttk.Frame(main)
        btn_row.pack(anchor="e")
        ttk.Button(btn_row, text="Cancel", command=on_cancel, width=10).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_row, text="OK", command=on_ok, width=10).pack(side=tk.LEFT)
        root.bind("<Return>", lambda _event: on_ok())
        root.bind("<Escape>", lambda _event: on_cancel())
        _center_tk_window(root)
        root.mainloop()
        return chosen[0]
    except Exception:
        suffix = "" if default is None else f" [{default}]"
        raw = input(f"Which aisle is this rack data for? ({'/'.join(str(a) for a in AISLE_NUMBERS)}){suffix}: ").strip()
        if not raw:
            return default
        try:
            return int(raw)
        except ValueError:
            return default


def _format_location_selection(sel: LocationSelection) -> str:
    if sel == "ALL":
        return "ALL"
    return ", ".join(sorted(sel))


def _filter_by_locations(all_rows: list[dict], sel: LocationSelection) -> list[dict]:
    if sel == "ALL":
        return list(all_rows)
    return [r for r in all_rows if r["location"] in sel]


def _analysis_rows_alt_aisles_side_30_only(rows: list[dict]) -> list[dict]:
    out: list[dict] = []
    for r in rows:
        if r["location"] in CD_SIDE_ALT_AISLES:
            if r["side"] == CD_SIDE_ALT_SUFFIX:
                out.append(r)
        else:
            out.append(r)
    return out


def _row_omit_fork_unreachable_columns(r: dict) -> bool:
    """True for columns no fork can validly pair-compare (01/02/71/72), on any side.

    The analysis metric is Fork2 − Fork1; if either fork can't reach the column its value
    is a zero placeholder, so the delta is meaningless — drop it regardless of CD/Non-CD.
    """
    return r["col"] in FORK_UNREACHABLE_COL_CODES


def _analysis_rows_for_workbook(filtered_rows: list[dict], even: bool) -> tuple[list[dict], int]:
    step1 = _analysis_rows_alt_aisles_side_30_only(filtered_rows)
    out: list[dict] = []
    dropped_nt = 0
    for r in step1:
        if _row_is_inhibited(r) or _row_is_nonexistent(r, even):
            continue
        if _row_omit_fork_unreachable_columns(r):
            dropped_nt += 1
            continue
        out.append(r)
    return out, dropped_nt


def _parse_location_cli(s: str, valid: set[str]) -> LocationSelection:
    u = s.strip().upper()
    if u == "ALL":
        return "ALL"
    parts: list[str] = []
    for raw in u.replace(";", ",").split(","):
        raw = raw.strip()
        if not raw:
            continue
        for token in raw.split():
            token = token.strip().upper()
            if len(token) == 2 and token.isdigit():
                parts.append(token)
    if not parts and len(u) == 2 and u.isdigit():
        parts = [u]
    fs = frozenset(parts)
    if not fs:
        raise ValueError("No valid two-digit location codes parsed.")
    bad = fs - valid
    if bad:
        raise ValueError(f"Unknown location(s): {sorted(bad)}. Valid: {sorted(valid)}")
    return fs


def _pick_location_filter(locations: list[str]) -> LocationSelection | None:
    locs_sorted = sorted(locations)
    valid = set(locs_sorted)

    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.title("Rack analysis — choose location(s)")
        root.attributes("-topmost", True)

        tk.Label(
            root,
            text=(
                "Select one or more locations (Ctrl+click or Shift+click).\n"
                "Or click 'All locations' to include every location in the analysis tabs.\n"
                "(All_Rack_Data = raw JSON; All_Rack_Data_Offsets when offsets applied.)"
            ),
            justify="left",
            padx=10,
            pady=8,
        ).pack()

        lb = tk.Listbox(
            root,
            height=min(18, len(locs_sorted) + 2),
            width=30,
            exportselection=False,
            selectmode=tk.EXTENDED,
        )
        for loc in locs_sorted:
            lb.insert(tk.END, loc)
        lb.pack(padx=10, pady=4)

        chosen: list[LocationSelection | None] = [None]

        def on_all() -> None:
            chosen[0] = "ALL"
            root.destroy()

        def on_ok() -> None:
            sel = lb.curselection()
            if not sel:
                messagebox.showwarning(
                    "Select location(s)",
                    "Select at least one location, or use All locations.",
                    parent=root,
                )
                return
            picked = frozenset(locs_sorted[i] for i in sel)
            chosen[0] = picked
            root.destroy()

        def on_cancel() -> None:
            chosen[0] = None
            root.destroy()

        root.protocol("WM_DELETE_WINDOW", on_cancel)

        bf = tk.Frame(root)
        bf.pack(pady=8)
        tk.Button(bf, text="All locations", width=14, command=on_all).pack(side=tk.LEFT, padx=4)
        tk.Button(bf, text="OK", width=10, command=on_ok).pack(side=tk.LEFT, padx=4)
        tk.Button(bf, text="Cancel", width=10, command=on_cancel).pack(side=tk.LEFT, padx=4)

        root.mainloop()

        if chosen[0] is not None:
            return chosen[0]
    except Exception:
        pass

    while True:
        typed = input("Location selection: ").strip()
        if not typed:
            continue
        try:
            return _parse_location_cli(typed, valid)
        except ValueError:
            continue


def _pstdev(vals: list[float]) -> float:
    return statistics.pstdev(vals) if len(vals) > 1 else 0.0


def _write_row(ws: Any, r: int, values: list[Any]) -> None:
    for c, v in enumerate(values, start=1):
        ws.cell(row=r, column=c, value=_cell_value_for_excel(v))


# ======================================================================================
# Step 3a — Analysis tabs
# ======================================================================================


def _sides_for_location(loc: str) -> list[tuple[str, str]]:
    if loc in CD_SIDE_ALT_AISLES:
        return [("30", "CD")]
    return [("10", "Non-CD"), ("40", "CD")]


def _levels_cols_for_subset(rows_subset: list[dict]) -> tuple[list[str], list[str]]:
    # levels DESCENDING so level 16 tops the matrix
    if not rows_subset:
        return [], []
    return _sort_levels(r["level"] for r in rows_subset), sorted({r["col"] for r in rows_subset})


def _write_matrix_block_at(
    ws: Any,
    start_row: int,
    block_title: str,
    rows_subset: list[dict],
    value_key: str,
    spec: float,
) -> tuple[int, str | None]:
    # cells over ±spec → red, else orange(low)→green(high)
    r = start_row
    levels, cols = _levels_cols_for_subset(rows_subset)
    _title(ws, block_title, r, kind="section", span=(1 + len(cols)) if cols else 1)
    r += 1
    if not rows_subset or not levels or not cols:
        ws.cell(
            row=r,
            column=1,
            value=_cell_value_for_excel(
                "No rows in the current location filter for this Location + Side combination."
            ),
        )
        r += 1
        return r + 1, None

    grid: dict[tuple[str, str], float] = {}
    for lev in levels:
        for c in cols:
            cell_rs = [x for x in rows_subset if x["level"] == lev and x["col"] == c]
            if cell_rs:
                grid[(lev, c)] = statistics.fmean(x[value_key] for x in cell_rs)
    vmin, vmax = _scale_bounds([v for v in grid.values() if abs(v) <= spec])

    def cell_fn(lev: str, c: str) -> MatrixCell:
        v = grid.get((lev, c))
        if v is None:
            return MatrixCell(value=EM_DASH)
        return MatrixCell(round(v, 2), _scale_fill(v, vmin, vmax, abs(v) > spec))

    r, first_data_row = _write_matrix(
        ws, r, levels, cols, cell_fn, row_label_fn=_level_label, col_label_fn=str
    )
    return r + 1, f"B{first_data_row}"


def _fill_combined_location_side_matrices(
    ws: Any,
    matrix_rows: list[dict],
    value_key: str,
    math_line: str,
    sheet_heading: str,
    sheet_subheading: str,
    spec: float,
) -> None:
    banner_span = 1 + len({x["col"] for x in matrix_rows}) if matrix_rows else 1
    _title(ws, sheet_heading, 1, kind="title", span=banner_span)
    _title(ws, sheet_subheading, 2, kind="subtitle", span=banner_span)
    _title(ws, math_line, 3, kind="subtitle", span=banner_span)
    r = 5
    freeze_cell: str | None = None
    for loc in ("01", "02", "03", "04", "05"):
        if not any(x["location"] == loc for x in matrix_rows):
            continue
        loc_name = LOCATION_NAMES.get(loc, loc)
        for side_code, side_role in _sides_for_location(loc):
            subset = [x for x in matrix_rows if x["location"] == loc and x["side"] == side_code]
            block_title = f"Location {loc} — {loc_name} — {side_role} (side {side_code})"
            nxt, fz = _write_matrix_block_at(ws, r, block_title, subset, value_key, spec)
            r = nxt
            if fz is not None and freeze_cell is None:
                freeze_cell = fz
    _finish_sheet(ws, freeze_cell or "A1")
    _set_matrix_widths(ws, 72)


# ======================================================================================
# Fork1_RackData_Analysis / Fork2_RackData_Analysis — raw (uncalculated) RTA values
# ======================================================================================

# (tab metric label, row-dict field for Fork1, row-dict field for Fork2, gradient scale axis,
#  (Fork1 raw-Pos-field, Fork2 raw-Pos-field) for the untaught-sentinel check, or None)
#
# The gradient is scoped to whichever axis is physically expected to cluster, so an outlier
# actually stands out instead of being washed out by the full block's range:
#   - Travel (X) is expected to hold steady down a COLUMN (same column, every level) — scale
#     per column.
#   - Hoist (Z) is expected to hold steady across a ROW (same level, every column) — scale
#     per row.
#   - Pos (Y) has no such directional expectation — one gradient across the whole block.
_FORK_RAW_METRICS: tuple[tuple[str, str, str, str, tuple[str, str] | None], ...] = (
    ("Travel (X)", "fork1_tp", "fork2_tp", "column", None),
    ("Hoist (Z)", "fork1_hp", "fork2_hp", "row", None),
    ("Pos (Y)", "fork1_pp", "fork2_pp", "block", ("fork1_pp_raw", "fork2_pp_raw")),
)

RTA_LOCATION = "03"

# A raw (pre-offset) Fork*Pos of exactly 1 is a teach-system sentinel meaning the position was
# not completely taught -- it is not a real measurement. Offsets shift this away from 1 (a raw
# 1 becomes 1 + offset), so the check must run on the untouched JSON value, never the
# offset-corrected one used for display/other shading.
UNTAUGHT_POS_RAW_VALUE = 1.0


def _is_untaught_pos_raw(v: float | None) -> bool:
    return v is not None and abs(v - UNTAUGHT_POS_RAW_VALUE) < 1e-6


def _rta_rows_for_fork_tabs(
    filtered_rows: list[dict], raw_rows: list[dict], even: bool
) -> list[dict]:
    # Unlike _analysis_rows_for_workbook, the fork-unreachable-column drop is NOT applied here:
    # each fork has its own (different) unreachable columns, applied per-fork when each tab is
    # built instead.
    raw_pos_by_key = {r["key"]: (r["fork1_pp"], r["fork2_pp"]) for r in raw_rows}
    rta_rows = [r for r in filtered_rows if r["location"] == RTA_LOCATION]
    kept = [r for r in rta_rows if not _row_is_inhibited(r) and not _row_is_nonexistent(r, even)]
    for r in kept:
        raw_f1, raw_f2 = raw_pos_by_key.get(r["key"], (None, None))
        r["fork1_pp_raw"] = raw_f1
        r["fork2_pp_raw"] = raw_f2
    return kept


def _write_raw_matrix_block_at(
    ws: Any,
    start_row: int,
    block_title: str,
    rows_subset: list[dict],
    value_key: str,
    scale_axis: str = "block",
    ignore_key: str | None = None,
) -> tuple[int, str | None]:
    r = start_row
    levels, cols = _levels_cols_for_subset(rows_subset)
    _title(ws, block_title, r, kind="section", span=(1 + len(cols)) if cols else 1)
    r += 1
    if not rows_subset or not levels or not cols:
        ws.cell(
            row=r,
            column=1,
            value=_cell_value_for_excel(
                "No rows in the current location filter for this Location + Side combination."
            ),
        )
        r += 1
        return r + 1, None

    grid: dict[tuple[str, str], float] = {}
    untaught: set[tuple[str, str]] = set()
    for lev in levels:
        for c in cols:
            cell_rs = [x for x in rows_subset if x["level"] == lev and x["col"] == c]
            if cell_rs:
                grid[(lev, c)] = statistics.fmean(x[value_key] for x in cell_rs)
                if ignore_key is not None and any(_is_untaught_pos_raw(x.get(ignore_key)) for x in cell_rs):
                    untaught.add((lev, c))

    scored = {k: v for k, v in grid.items() if k not in untaught}
    if scale_axis == "column":
        bounds = {c: _scale_bounds([scored[(lev, c)] for lev in levels if (lev, c) in scored]) for c in cols}
        bounds_of = lambda lev, c: bounds[c]
    elif scale_axis == "row":
        bounds = {lev: _scale_bounds([scored[(lev, c)] for c in cols if (lev, c) in scored]) for lev in levels}
        bounds_of = lambda lev, c: bounds[lev]
    else:
        block_bounds = _scale_bounds(list(scored.values()))
        bounds_of = lambda lev, c: block_bounds

    def cell_fn(lev: str, c: str) -> MatrixCell:
        v = grid.get((lev, c))
        if v is None:
            return MatrixCell(value=EM_DASH)
        if (lev, c) in untaught:
            return MatrixCell(round(v, 2))
        vmin, vmax = bounds_of(lev, c)
        return MatrixCell(round(v, 2), _scale_fill(v, vmin, vmax, False))

    r, first_data_row = _write_matrix(
        ws, r, levels, cols, cell_fn, row_label_fn=_level_label, col_label_fn=str
    )
    return r + 1, f"B{first_data_row}"


def _write_fork_raw_tab(
    wb: Any,
    sheet_name: str,
    fork_label: str,
    rta_rows: list[dict],
    fork_field_index: int,
    unreachable_cols: frozenset[int],
    unreachable_note: str,
) -> None:
    ws = wb.create_sheet(sheet_name)
    banner_span = 1 + len({x["col"] for x in rta_rows}) if rta_rows else 1
    _title(
        ws,
        f"{fork_label} raw rack data — RTA (Location 03)",
        1,
        kind="title",
        span=banner_span,
    )
    _title(
        ws,
        f"Raw {fork_label} Travel / Hoist / Pos values (no calculation), Level × Column, "
        f"Non-CD (10) and CD (40) sides. Offset-corrected data. {unreachable_note}",
        2,
        kind="subtitle",
        span=banner_span,
    )
    _title(
        ws,
        "Gradient is scoped per metric: Travel shades within each column (values expected to "
        "hold steady down a column), Hoist shades within each row (values expected to hold "
        "steady across a row), Pos shades across the whole block. No pass/fail colour. "
        "Pos cells where the raw (pre-offset) value is exactly 1 — not completely taught — "
        "are shown unshaded and excluded from the gradient bounds.",
        3,
        kind="subtitle",
        span=banner_span,
    )
    r = 5
    freeze_cell: str | None = None
    for side_code, side_role in _sides_for_location(RTA_LOCATION):
        subset_all = [x for x in rta_rows if x["side"] == side_code]
        subset = [x for x in subset_all if int(x["col"]) not in unreachable_cols]
        block_title = f"Location 03 — RTA — {side_role} (side {side_code})"
        for metric_label, f1_key, f2_key, scale_axis, raw_pos_keys in _FORK_RAW_METRICS:
            value_key = f1_key if fork_field_index == 1 else f2_key
            ignore_key = None if raw_pos_keys is None else raw_pos_keys[fork_field_index - 1]
            nxt, fz = _write_raw_matrix_block_at(
                ws, r, f"{block_title} — {fork_label} {metric_label}", subset, value_key,
                scale_axis, ignore_key,
            )
            r = nxt
            if fz is not None and freeze_cell is None:
                freeze_cell = fz
    _finish_sheet(ws, freeze_cell or "A1")
    _set_matrix_widths(ws, 72)


# ======================================================================================
# Step 1 — Parse the JSON into rows
# ======================================================================================


def _parse_rack_rows(
    data: dict[str, Any],
    offsets: LaserPaddleOffsets | None = None,
) -> tuple[list[dict], int, int]:
    rows: list[dict] = []
    skipped_bad_key = 0
    skipped_missing = 0

    for key, e in data.items():
        if not isinstance(e, dict):
            skipped_bad_key += 1
            continue
        if len(key) != 8 or not key.isdigit():
            skipped_bad_key += 1
            continue
        if not all(k in e for k in REQUIRED_KEYS):
            skipped_missing += 1
            continue

        level = key[2:4]
        col = key[4:6]
        side = key[6:8]
        location = key[:2]

        f1t = float(e["Fork1TravelPos"])
        f2t = float(e["Fork2TravelPos"])
        f1h = float(e["Fork1HoistPos"])
        f2h = float(e["Fork2HoistPos"])
        f1p = float(e["Fork1Pos"])
        f2p = float(e["Fork2Pos"])
        if offsets is not None:
            f1t, f2t, f1h, f2h, f1p, f2p = _apply_laser_paddle_offsets(
                location, side, f1t, f2t, f1h, f2h, f1p, f2p, offsets
            )
        f2_expect = f1t + TRAVEL_F2_AHEAD_OF_F1_MM
        d_h = f2h - f1h
        d_p = f2p - f1p
        d_t = f2t - f1t
        d_te = f2t - f2_expect

        rows.append(
            {
                "key": key,
                "location": location,
                "level": level,
                "col": col,
                "side": side,
                "fork1_tp": f1t,
                "fork1_hp": f1h,
                "fork1_pp": f1p,
                "fork2_tp": f2t,
                "fork2_hp": f2h,
                "fork2_pp": f2p,
                "d_travel": d_t,
                "d_travel_err": d_te,
                "d_hoist": d_h,
                "d_pos": d_p,
                "inhibit": e.get("Inhibit", ""),
            }
        )

    rows.sort(key=lambda r: r["key"])
    return rows, skipped_bad_key, skipped_missing


# ======================================================================================
# Step 3b — Expected rack layout
# ======================================================================================
# Ported verbatim from AutoTeach_Log_Consolidation.py — keep the two copies in sync.

MAX_LEVEL = 16
MAX_COLUMN = 72


def expected_area(line: str, level: int, col: int, even: bool = False) -> str | None:
    """Returns 'RTA', 'HTA', or None (empty / no rack).

    Non-CD / Aging (line 10): all RTA. Odd and even aisles differ at two column blocks.
        cols 1-4  exist L3-13   cols 5-8  exist L3-15
        cols 9-68 exist L1-16   cols 69-72 exist L2-16
        EVEN aisles only: cols 21-24 exist L3-16 (levels 1 and 2 removed) and
        cols 69-72 exist L1-16 (level 1 added).
    CD (line 40): HTA/RTA schema differs by aisle parity (cols 1-20 are HTA either way).
      Odd:  1-32 HTA, 33-48 RTA, 49-60 HTA, 61-68 RTA, 69-72 RTA (L2-16).
      Even: 1-20 HTA, 21-24 RTA (L2-16), 25-32 RTA, 33-44 HTA, 45-60 RTA, 61-72 HTA.
    """
    if line == "10":
        if 1 <= col <= 4:
            return "RTA" if 3 <= level <= 13 else None
        if 5 <= col <= 8:
            return "RTA" if 3 <= level <= 15 else None
        if 69 <= col <= 72:
            # Even aisles gained a level 1 here, so the block runs the full L1-16.
            # Odd aisles still have no rack at level 1.
            return "RTA" if (even or level >= 2) else None
        if 9 <= col <= 68:
            # Even aisles have no rack at columns 21-24 levels 1 AND 2; both were removed.
            # Odd aisles keep the full L1-16 height here.
            if even and 21 <= col <= 24:
                return "RTA" if level >= 3 else None
            return "RTA"
        return None
    if line == "40":
        if even:
            if 1 <= col <= 20:
                return "HTA"
            if 21 <= col <= 24:
                return "RTA" if level >= 2 else None
            if 25 <= col <= 32:
                return "RTA"
            if 33 <= col <= 44:
                return "HTA"
            if 45 <= col <= 60:
                return "RTA"
            if 61 <= col <= 72:
                return "HTA"
            return None
        if 1 <= col <= 32:
            return "HTA"
        if 33 <= col <= 48:
            return "RTA"
        if 49 <= col <= 60:
            return "HTA"
        if 61 <= col <= 68:
            return "RTA"
        if 69 <= col <= 72:
            return "RTA" if level >= 2 else None
        return None
    return None


def _row_is_nonexistent(r: dict, even: bool) -> bool:
    # A cell is kept if it is RTA or HTA — only expected_area()==None is dropped, so a row's
    # real aging-side (line 10) data is never removed for lacking an HTA label.
    if r["location"] not in ("03", "04") or r["side"] not in ("10", "40"):
        return False
    return expected_area(r["side"], int(r["level"]), int(r["col"]), even) is None


# ======================================================================================
# Step 3c — RTA/HTA fixture checks
# ======================================================================================

# Fork 1 cannot reach columns 1-2; Fork 2 cannot reach columns 71-72 (zero-filled in JSON).
FORK1_UNREACHABLE_COLS = frozenset({1, 2})
FORK2_UNREACHABLE_COLS = frozenset({71, 72})
FIXTURE_XYZ_SKIP_COLS = FORK1_UNREACHABLE_COLS | FORK2_UNREACHABLE_COLS

FIXTURE_TOLERANCES: dict[str, tuple[int | float, str]] = {
    "Y_DIFF": (8, "mm"),
    "Z_DIFF": (6, "mm"),
    "X_DIFF": (8, "mm"),
    "GANGTRAVEL": (4, "mm"),
    "GANGHOIST": (6, "mm"),
    "LOADBEAM": (7, "mm"),
    "HOISTROW": (10, "mm"),
    "BAY_TRAVEL": (10, "mm"),
    "BAY_POS": (5, "mm"),
}


def _rack_rows_to_fixture_df(all_rows: list[dict], even: bool) -> Any | None:
    try:
        import pandas as pd
    except ImportError:
        return None
    recs: list[dict[str, Any]] = []
    for r in all_rows:
        loc = r["location"]
        if loc not in ("03", "04"):
            continue
        side = r["side"]
        if side not in ("10", "40"):
            continue
        if _row_is_inhibited(r) or _row_is_nonexistent(r, even):
            continue
        recs.append(
            {
                "KEY": r["key"],
                "LOCATION": int(loc),
                "ROW": int(r["level"]),
                "COLUMN": int(r["col"]),
                "LINE": int(side),
                "FORK1POS": float(r["fork1_pp"]),
                "FORK2POS": float(r["fork2_pp"]),
                "FORK1HOISTPOS": float(r["fork1_hp"]),
                "FORK2HOISTPOS": float(r["fork2_hp"]),
                "FORK1TRAVELPOS": float(r["fork1_tp"]),
                "FORK2TRAVELPOS": float(r["fork2_tp"]),
            }
        )
    return pd.DataFrame(recs) if recs else pd.DataFrame()


def _fixture_group_of(hi: dict[str, int], row: list[Any]) -> tuple[str, str, str]:
    # Area  ← "Area" (RTA/HTA) or "Location" (3→RTA, 4→HTA).
    # Side  ← "Side" (Aging→Non-CD, CD) or "Line" (10→Non-CD, 40→CD).
    # Fork  ← "Fork" (→ "Fork1"/"Fork2"), else "" when the check is a fork pair.
    if "Area" in hi:
        area = str(row[hi["Area"]])
    elif "Location" in hi:
        area = "RTA" if int(row[hi["Location"]]) == 3 else "HTA"
    else:
        area = ""
    if "Side" in hi:
        side = "CD" if str(row[hi["Side"]]) == "CD" else "Non-CD"
    elif "Line" in hi:
        side = "CD" if int(row[hi["Line"]]) == 40 else "Non-CD"
    else:
        side = ""
    fork = f"Fork{row[hi['Fork']]}" if "Fork" in hi else ""
    return area, side, fork


def _fixture_group_order(key: tuple[str, str, str]) -> tuple[int, int, str]:
    area, side, fork = key
    return (0 if area == "RTA" else 1, 0 if side == "Non-CD" else 1, fork)


def _count_fixture_by_area_side(headers: list[str], rows: list[list[Any]]) -> dict[tuple[str, str], int]:
    hi = {h: i for i, h in enumerate(headers)}
    counts: dict[tuple[str, str], int] = defaultdict(int)
    for row in rows:
        area, side, _ = _fixture_group_of(hi, row)
        counts[(area, side)] += 1
    return counts


def _order_fixture_rows(headers: list[str], rows: list[list[Any]]) -> list[list[Any]]:
    # The level column is called "Row" in the fixture tables and "Level" elsewhere; either
    # name is honoured.
    hi = {h: i for i, h in enumerate(headers)}
    lvl_i = hi.get("Row", hi.get("Level"))
    if lvl_i is None:
        return list(rows)
    col_i = hi.get("Column", hi.get("Col"))
    tiebreak = (lambda row: row[col_i]) if col_i is not None else None
    return _sort_level_rows(rows, lambda row: row[lvl_i], tiebreak)


def _write_grouped_fixture_section(
    ws: Any,
    start_row: int,
    section_title: str,
    subtitle: str,
    headers: list[str],
    rows: list[list[Any]],
) -> int:
    width = len(headers)
    r = start_row
    _title(ws, section_title, r, kind="section", span=width)
    r += 1
    if subtitle:
        _title(ws, subtitle, r, kind="subtitle", span=width)
        r += 1
    if not rows:
        _title(ws, "No failures.", r, kind="subtitle", span=width)
        return r + 2

    hi = {h: i for i, h in enumerate(headers)}
    groups: dict[tuple[str, str, str], list[list[Any]]] = defaultdict(list)
    for row in rows:
        groups[_fixture_group_of(hi, row)].append(row)

    s = _styles()
    for key in sorted(groups, key=_fixture_group_order):
        label = " · ".join(p for p in key if p) or "All"
        _title(ws, f"{label}  —  {len(groups[key])} failed", r, kind="section", span=width)
        r += 1
        r = _write_styled_table(
            ws, r, headers, _order_fixture_rows(headers, groups[key]), fill=s["fill_fail"]
        )
        r += 1
    return r + 1


GangBucketKey = tuple[int, int]
GangBucketStats = dict[str, Any]


def _gang_side_name(line: int) -> str:
    return "Aging" if line == 10 else "CD"


def _gang_bucket_label(location: int, line: int) -> str:
    area = "RTA" if location == 3 else "HTA"
    return f"{area} — {_gang_side_name(line)} (line {line})"


def _new_gang_buckets() -> dict[GangBucketKey, GangBucketStats]:
    buckets: dict[GangBucketKey, GangBucketStats] = {}
    for location in (3, 4):
        for line in (10, 40):
            buckets[(location, line)] = {
                "label": _gang_bucket_label(location, line),
                "pairs": 0,
                "failed": 0,
                "diffs": [],
            }
    return buckets


def _record_gang_pair(
    buckets: dict[GangBucketKey, GangBucketStats],
    location: int,
    line: int,
    diff: float,
    tolerance: float,
) -> bool:
    bucket = buckets[(location, line)]
    bucket["pairs"] += 1
    bucket["diffs"].append(diff)
    failed = diff > tolerance
    if failed:
        bucket["failed"] += 1
    return failed


def _fixture_check_y(df: Any) -> tuple[str, str, list[str], list[list[Any]], int, int]:
    failures: list[list[Any]] = []
    failed_y_rta = 0
    failed_y_hta = 0
    tolerance, unit = FIXTURE_TOLERANCES["Y_DIFF"]
    for _idx, row_data in df.iterrows():
        location = int(row_data["LOCATION"])
        row = int(row_data["ROW"])
        column = int(row_data["COLUMN"])
        line = int(row_data["LINE"])
        if column in FIXTURE_XYZ_SKIP_COLS:
            continue
        y_diff = abs(float(row_data["FORK1POS"]) - float(row_data["FORK2POS"]))
        if y_diff > tolerance:
            if location == 3:
                failed_y_rta += 1
            elif location == 4:
                failed_y_hta += 1
            if location in (3, 4) and line in (10, 40):
                key = str(row_data.get("KEY", "")) or f"{location:02d}{row:02d}{column:02d}{line:02d}"
                msg = (
                    f"{key} is not good: The Y difference is {y_diff:.2f} {unit} "
                    f"and the spec is {tolerance} {unit}."
                )
                failures.append([key, location, row, column, line, round(y_diff, 4), tolerance, msg])
    title = "Y difference (|Fork1Pos − Fork2Pos|)"
    subtitle = f"Fail if |F1 Y − F2 Y| > {tolerance} {unit}. RTA 03 / HTA 04. Columns 1/2/71/72 excluded."
    headers = ["Key", "Location", "Row", "Column", "Line", "Y_diff_mm", "Spec_mm", "Message"]
    return title, subtitle, headers, failures, failed_y_rta, failed_y_hta


def _fixture_check_z(df: Any) -> tuple[str, str, list[str], list[list[Any]], int, int]:
    failures: list[list[Any]] = []
    failed_z_rta = 0
    failed_z_hta = 0
    tolerance, unit = FIXTURE_TOLERANCES["Z_DIFF"]
    for _idx, row_data in df.iterrows():
        location = int(row_data["LOCATION"])
        row = int(row_data["ROW"])
        column = int(row_data["COLUMN"])
        line = int(row_data["LINE"])
        if column in FIXTURE_XYZ_SKIP_COLS:
            continue
        z_diff = abs(float(row_data["FORK1HOISTPOS"]) - float(row_data["FORK2HOISTPOS"]))
        if z_diff > tolerance:
            if location == 3:
                failed_z_rta += 1
            elif location == 4:
                failed_z_hta += 1
            if location in (3, 4) and line in (10, 40):
                key = str(row_data.get("KEY", "")) or f"{location:02d}{row:02d}{column:02d}{line:02d}"
                msg = (
                    f"{key} is not good: The Z difference is {z_diff:.2f} {unit} "
                    f"and the spec is {tolerance} {unit}."
                )
                failures.append([key, location, row, column, line, round(z_diff, 4), tolerance, msg])
    title = "Z difference (|Fork1HoistPos − Fork2HoistPos|)"
    subtitle = f"Fail if |F1 Z − F2 Z| > {tolerance} {unit}. RTA 03 / HTA 04. Columns 1/2/71/72 excluded."
    headers = ["Key", "Location", "Row", "Column", "Line", "Z_diff_mm", "Spec_mm", "Message"]
    return title, subtitle, headers, failures, failed_z_rta, failed_z_hta


def _fixture_check_x(df: Any) -> tuple[str, str, list[str], list[list[Any]], int, int]:
    failures: list[list[Any]] = []
    failed_x_rta = 0
    failed_x_hta = 0
    tolerance, unit = FIXTURE_TOLERANCES["X_DIFF"]
    for _idx, row_data in df.iterrows():
        location = int(row_data["LOCATION"])
        row = int(row_data["ROW"])
        column = int(row_data["COLUMN"])
        line = int(row_data["LINE"])
        if column in FIXTURE_XYZ_SKIP_COLS:
            continue
        x_diff = abs(float(row_data["FORK1TRAVELPOS"]) - float(row_data["FORK2TRAVELPOS"]))
        if x_diff < (TRAVEL_TARGET - tolerance) or x_diff > (TRAVEL_TARGET + tolerance):
            if location == 3:
                failed_x_rta += 1
            elif location == 4:
                failed_x_hta += 1
            if location in (3, 4) and line in (10, 40):
                key = str(row_data.get("KEY", "")) or f"{location:02d}{row:02d}{column:02d}{line:02d}"
                dev = abs(TRAVEL_TARGET - x_diff)
                msg = (
                    f"{key} is not good: deviation from nominal {int(TRAVEL_TARGET)} {unit} is "
                    f"{dev:.2f} {unit} (spec ±{tolerance} {unit})."
                )
                failures.append([key, location, row, column, line, round(x_diff, 4), tolerance, round(dev, 4), msg])
    title = "X / travel separation (|F1Travel − F2Travel| ~ 1180 mm)"
    subtitle = (
        f"Nominal |F1 X − F2 X| = {int(TRAVEL_TARGET)} {unit} ± {tolerance} {unit}. "
        "RTA 03 / HTA 04. Columns 1/2/71/72 excluded."
    )
    headers = [
        "Key",
        "Location",
        "Row",
        "Column",
        "Line",
        "X_separation_mm",
        "Spec_mm",
        "Deviation_from_nom_mm",
        "Message",
    ]
    return title, subtitle, headers, failures, failed_x_rta, failed_x_hta


def _fixture_max_col_by_loc_line(df: Any) -> dict[tuple[int, int], Any]:
    max_col_by_location_line: dict[tuple[int, int], Any] = {}
    for location in (3, 4):
        for line in (10, 40):
            filtered = df[(df["LOCATION"] == location) & (df["LINE"] == line)]
            if not filtered.empty:
                max_col = int(filtered["COLUMN"].max())
                max_col_by_location_line[(location, line)] = range(1, max_col + 1)
            else:
                max_col_by_location_line[(location, line)] = range(1, 5)
    return max_col_by_location_line


def _fixture_check_gang_combined(
    df: Any,
    mode: str,
    max_row: int,
    max_col_by_location_line: dict[tuple[int, int], Any],
) -> tuple[str, str, list[str], list[list[Any]], int, int, int, int, dict[GangBucketKey, GangBucketStats], list[dict]]:
    failures: list[list[Any]] = []
    cells: list[dict] = []
    buckets = _new_gang_buckets()
    tol_key = "GANG" + mode.upper()
    tolerance, unit = FIXTURE_TOLERANCES[tol_key]
    area_name = {3: "RTA", 4: "HTA"}

    for location in (3, 4):
        for line in (10, 40):
            col_range = max_col_by_location_line.get((location, line), range(1, 2))
            for row in range(1, max_row + 1):
                row_data = df[
                    (df["LOCATION"] == location) & (df["LINE"] == line) & (df["ROW"] == row)
                ]
                for col_start in col_range[::4]:
                    for col1, col2 in ((col_start + 2, col_start), (col_start + 3, col_start + 1)):
                        if col1 > col_range[-1] or col2 > col_range[-1]:
                            continue
                        fork1 = row_data[row_data["COLUMN"] == col1]
                        fork2 = row_data[row_data["COLUMN"] == col2]
                        if fork1.empty or fork2.empty:
                            continue
                        # Fork 1 reads col1, Fork 2 reads col2 — skip if either is unreachable.
                        if col1 in FORK1_UNREACHABLE_COLS or col2 in FORK2_UNREACHABLE_COLS:
                            continue
                        if mode.upper() == "TRAVEL":
                            diff = abs(
                                float(fork1["FORK1TRAVELPOS"].values[0])
                                - float(fork2["FORK2TRAVELPOS"].values[0])
                            )
                        else:
                            diff = abs(
                                float(fork1["FORK1HOISTPOS"].values[0])
                                - float(fork2["FORK2HOISTPOS"].values[0])
                            )

                        failed = _record_gang_pair(buckets, location, line, diff, tolerance)
                        cells.append(
                            {
                                "loc": location, "line": line, "fork": 0,
                                "rlabel": f"L{row}", "rsort": row,
                                "clabel": f"{col1}↔{col2}", "csort": col1, "error": diff,
                            }
                        )
                        if failed:
                            ref = f"{location:02d}{row:02d}{col1:02d}{line:02d}"
                            side = _gang_side_name(line)
                            area = area_name[location]
                            failures.append(
                                [
                                    ref,
                                    area,
                                    side,
                                    row,
                                    col1,
                                    col2,
                                    line,
                                    round(diff, 4),
                                    tolerance,
                                ]
                            )

    failed_rta = buckets[(3, 10)]["failed"] + buckets[(3, 40)]["failed"]
    failed_hta = buckets[(4, 10)]["failed"] + buckets[(4, 40)]["failed"]
    total_pairs_rta = buckets[(3, 10)]["pairs"] + buckets[(3, 40)]["pairs"]
    total_pairs_hta = buckets[(4, 10)]["pairs"] + buckets[(4, 40)]["pairs"]

    fail_headers = [
        "Ref_key",
        "Area",
        "Side",
        "Row",
        "Col1",
        "Col2",
        "Line",
        "Diff_mm",
        "Spec_mm",
    ]
    if mode.upper() == "TRAVEL":
        title = "Gang pick/place — travel (|F1Travel@col1 − F2Travel@col2|)"
        subtitle = (
            f"Tolerance {tolerance} {unit}. Pairs (col_start+2 vs col_start), "
            f"(col_start+3 vs col_start+1) every 4 cols. Summary split by RTA/HTA and Aging (10) vs CD (40)."
        )
        return (
            title,
            subtitle,
            fail_headers,
            failures,
            failed_rta,
            failed_hta,
            total_pairs_rta,
            total_pairs_hta,
            buckets,
            cells,
        )
    title = "Gang pick/place — hoist (|F1Hoist@col1 − F2Hoist@col2|)"
    subtitle = (
        f"Tolerance {tolerance} {unit}. Same column-pair pattern as gang travel. "
        "Summary split by RTA/HTA and Aging (10) vs CD (40)."
    )
    return (
        title,
        subtitle,
        fail_headers,
        failures,
        failed_rta,
        failed_hta,
        total_pairs_rta,
        total_pairs_hta,
        buckets,
        cells,
    )


def _fixture_check_hoist_between_rows(df: Any, max_row: int) -> tuple[str, str, list[str], list[list[Any]], int, int, list[dict]]:
    failures: list[list[Any]] = []
    cells: list[dict] = []
    failed_rta = 0
    failed_hta = 0
    tolerance, unit = FIXTURE_TOLERANCES["HOISTROW"]
    for location in (3, 4):
        area = "RTA" if location == 3 else "HTA"
        for line in (10, 40):
            data = df[(df["LOCATION"] == location) & (df["LINE"] == line)]
            for col in sorted(data["COLUMN"].unique()):
                for row in range(1, max_row):
                    expected_spacing = 1150 if row % 2 == 1 else 1200
                    f1 = data[(data["ROW"] == row) & (data["COLUMN"] == col)]["FORK1HOISTPOS"]
                    f2 = data[(data["ROW"] == row + 1) & (data["COLUMN"] == col)]["FORK1HOISTPOS"]
                    f3 = data[(data["ROW"] == row) & (data["COLUMN"] == col)]["FORK2HOISTPOS"]
                    f4 = data[(data["ROW"] == row + 1) & (data["COLUMN"] == col)]["FORK2HOISTPOS"]
                    if f1.empty or f2.empty:
                        continue
                    diff1 = abs(float(f1.values[0]) + expected_spacing - float(f2.values[0]))
                    has_f34 = not f3.empty and not f4.empty
                    diff2 = (
                        abs(float(f3.values[0]) + expected_spacing - float(f4.values[0])) if has_f34 else None
                    )
                    rlabel = f"L{row}→L{row + 1}"
                    if col not in FORK1_UNREACHABLE_COLS:
                        cells.append({"loc": location, "line": line, "fork": 1, "rlabel": rlabel,
                                      "rsort": row, "clabel": f"C{col}", "csort": col, "error": diff1})
                    if diff2 is not None and col not in FORK2_UNREACHABLE_COLS:
                        cells.append({"loc": location, "line": line, "fork": 2, "rlabel": rlabel,
                                      "rsort": row, "clabel": f"C{col}", "csort": col, "error": diff2})
                    if diff1 > tolerance and col not in FORK1_UNREACHABLE_COLS:
                        if location == 3:
                            failed_rta += 1
                        else:
                            failed_hta += 1
                        failures.append(
                            [area, 1, location, row, row + 1, col, line, round(diff1, 4), tolerance, expected_spacing]
                        )
                    if diff2 is not None and diff2 > tolerance and col not in FORK2_UNREACHABLE_COLS:
                        if location == 3:
                            failed_rta += 1
                        else:
                            failed_hta += 1
                        failures.append(
                            [area, 2, location, row, row + 1, col, line, round(diff2, 4), tolerance, expected_spacing]
                        )
    title = "Hoist spacing between adjacent rows (Fork1 / Fork2)"
    subtitle = (
        f"Expected row-to-row hoist step 1150 mm (odd lower row index) or 1200 mm (even); "
        f"fail if deviation > {tolerance} {unit}."
    )
    headers = [
        "Area",
        "Fork",
        "Location",
        "Row_from",
        "Row_to",
        "Column",
        "Line",
        "Diff_mm",
        "Spec_mm",
        "Expected_spacing_mm",
    ]
    return title, subtitle, headers, failures, failed_rta, failed_hta, cells


# Column geometry inside the rack, in mm. Columns sit BAY_COLUMN_PITCH_MM apart inside a
# 4-column bay; stepping from a bay's last column to the next bay's first column is wider.
# Verified against taught data: within-bay 589.2-592.9, bay-to-bay 738.6-741.5.
BAY_COLUMN_PITCH_MM = 590
BAY_GAP_MM = 740

# The three bay-level checks differ only in which field they read and whether the four
# columns of a bay are meant to share that coordinate. Pos (Y) and Hoist (Z) must match
# across a bay, so the raw difference IS the error. Travel (X) must not -- the columns are
# a pitch apart -- so its nominal span is subtracted first; without that every cell would
# simply read ~590/1180/1770 and nothing could ever fail. Travel is also the only check
# that steps across a bay boundary (Pos/Hoist are deliberately within-bay only).
# (field suffix, FIXTURE_TOLERANCES key, display label, nominal pitch)
_BAY_PAIR_CHECKS: dict[str, tuple[str, str, str, int]] = {
    "POS": ("POS", "BAY_POS", "Pos (Y)", 0),
    "HOIST": ("HOISTPOS", "LOADBEAM", "Hoist (Z)", 0),
    "TRAVEL": ("TRAVELPOS", "BAY_TRAVEL", "Travel (X)", BAY_COLUMN_PITCH_MM),
}


def _fixture_check_bay_pairs(
    df: Any, max_row: int, max_col_by_location_line: dict[tuple[int, int], Any], kind: str
) -> tuple[str, str, list[str], list[list[Any]], int, int, list[dict]]:
    suffix, tol_key, label, pitch = _BAY_PAIR_CHECKS[kind]
    tolerance, unit = FIXTURE_TOLERANCES[tol_key]
    rows_rta: list[list[Any]] = []
    rows_hta: list[list[Any]] = []
    cells: list[dict] = []
    for location in (3, 4):
        area = "RTA" if location == 3 else "HTA"
        for line in (10, 40):
            col_range = max_col_by_location_line.get((location, line), range(1, 2))
            for row in range(1, max_row + 1):
                row_data = df[
                    (df["LOCATION"] == location) & (df["LINE"] == line) & (df["ROW"] == row)
                ]
                for col_start in col_range[::4]:
                    cols = [col_start + i for i in range(4)]
                    if any(c > col_range[-1] for c in cols):
                        continue
                    # Travel also spans the gap to the next bay's first column; that pair is
                    # the only one whose nominal is BAY_GAP_MM rather than a pitch multiple.
                    next_col = col_start + 4
                    span = cols + ([next_col] if next_col <= col_range[-1] else [])
                    f1: dict[int, float] = {}
                    f2: dict[int, float] = {}
                    for col in span:
                        val = row_data[row_data["COLUMN"] == col]
                        if not val.empty:
                            f1[col] = float(val["FORK1" + suffix].values[0])
                            f2[col] = float(val["FORK2" + suffix].values[0])
                    pairs = [
                        (cols[i], cols[j], pitch * (cols[j] - cols[i]))
                        for i in range(4)
                        for j in range(i + 1, 4)
                    ]
                    if pitch and next_col in f1 | f2:
                        pairs.append((cols[3], next_col, BAY_GAP_MM))
                    for c_a, c_b, nominal in pairs:
                        for fork, vals, unreach in (
                            (1, f1, FORK1_UNREACHABLE_COLS),
                            (2, f2, FORK2_UNREACHABLE_COLS),
                        ):
                            if c_a not in vals or c_b not in vals:
                                continue
                            if c_a in unreach or c_b in unreach:
                                continue
                            error = abs(abs(vals[c_a] - vals[c_b]) - nominal)
                            cells.append({"loc": location, "line": line, "fork": fork, "rlabel": f"L{row}",
                                          "rsort": row, "clabel": f"C{c_a}-C{c_b}", "csort": (c_a, c_b), "error": error})
                            if error > tolerance:
                                target = rows_rta if location == 3 else rows_hta
                                target.append([area, fork, location, row, c_a, c_b, line, round(error, 4), tolerance])
    rows_out = rows_rta + rows_hta
    title = f"Bay {label} level — all column pairs (Fork1 / Fork2)"
    if pitch:
        subtitle = (
            f"Within each 4-col bay, all 6 column pairs, plus col4 → next bay's col1. "
            f"Nominal = {pitch} mm × (col_b − col_a) within a bay and {BAY_GAP_MM} mm across "
            f"the bay gap; fail if |measured − nominal| > {tolerance} {unit}."
        )
    else:
        subtitle = (
            f"Within each 4-col bay, all 6 column pairs (no bay-to-bay comparison); "
            f"fail if the difference between the two columns exceeds {tolerance} {unit}."
        )
    headers = ["Area", "Fork", "Location", "Row", "Col_a", "Col_b", "Line", "Error_mm", "Spec_mm"]
    return title, subtitle, headers, rows_out, len(rows_rta), len(rows_hta), cells


def _build_summary_sheet(
    wb: Any,
    checks: list[tuple[str, list[str], list[list[Any]]]],
    input_rows: int,
    max_row: int,
    aisle: int,
    even: bool,
) -> None:
    s = _styles()
    ws = wb.create_sheet("Summary")
    headers = ["Check", "RTA Non-CD", "RTA CD", "HTA Non-CD", "HTA CD", "Total failed", "Status"]
    _title(ws, "RTA / HTA fixture check summary", 1, kind="title", span=len(headers))
    _title(
        ws,
        f"Aisle {aisle} ({'EVEN' if even else 'ODD'} CD layout).  "
        f"Locations 03 (RTA) / 04 (HTA), sides 10 (Non-CD) / 40 (CD).  Offset-corrected data.  "
        f"Fixture input rows: {input_rows}.  Max level: {max_row}.  "
        f"Green = passed, orange = failures found.",
        2,
        kind="subtitle",
        span=len(headers),
    )
    r = _style_header(ws, headers, row=4)
    area_side_cols = [("RTA", "Non-CD"), ("RTA", "CD"), ("HTA", "Non-CD"), ("HTA", "CD")]
    for name, hdrs, rows in checks:
        counts = _count_fixture_by_area_side(hdrs, rows)
        vals = [counts.get(k, 0) for k in area_side_cols]
        total = sum(vals)
        line = [name, *vals, total, "PASS" if total == 0 else "FAIL"]
        for c, v in enumerate(line, start=1):
            cell = ws.cell(row=r, column=c, value=_cell_value_for_excel(v))
            cell.border = s["border"]
            if c > 1:
                cell.alignment = s["center"]
            if 2 <= c <= 6 and v:
                cell.fill = s["fill_fail"]
        ws.cell(row=r, column=len(line)).fill = s["fill_ok"] if total == 0 else s["fill_fail"]
        r += 1
    _set_widths(ws, [16, 12, 10, 12, 10, 13, 9])
    _finish_sheet(ws, "B5")
    wb.move_sheet("Summary", offset=-wb.sheetnames.index("Summary"))


def _append_fixture_sheets_to_workbook(
    wb: Any, all_rows: list[dict], even: bool, aisle: int
) -> dict[str, Any]:
    summary: dict[str, Any] = {"fixture_status": "ok"}
    df = _rack_rows_to_fixture_df(all_rows, even)
    if df is None or df.empty:
        summary["fixture_status"] = "skipped_no_pandas" if df is None else "no_03_04_line_10_40"
        ws = wb.create_sheet("Summary")
        _title(ws, "RTA / HTA fixture checks", 1, kind="title")
        note = (
            "Skipped — install pandas (pip install pandas)."
            if df is None
            else "No rows: need rack keys for locations 03/04 with side (LINE) 10 or 40."
        )
        _title(ws, note, 2, kind="subtitle")
        wb.move_sheet("Summary", offset=-wb.sheetnames.index("Summary"))
        return summary

    max_row = int(df["ROW"].max())
    max_col_map = _fixture_max_col_by_loc_line(df)

    t_y, s_y, h_y, r_y, fy_r, fy_h = _fixture_check_y(df)
    t_z, s_z, h_z, r_z, fz_r, fz_h = _fixture_check_z(df)
    t_x, s_x, h_x, r_x, fx_r, fx_h = _fixture_check_x(df)
    tg_t, sg_t, hg_t, rg_t, fgrt_r, fgrt_h, tprt, tpht, gtb, gang_travel_cells = (
        _fixture_check_gang_combined(df, "TRAVEL", max_row, max_col_map)
    )
    tg_h, sg_h, hg_h, rg_h, fgrh_r, fgrh_h, tphr, tphh, ghb, gang_hoist_cells = (
        _fixture_check_gang_combined(df, "HOIST", max_row, max_col_map)
    )
    gang_travel_tol, _ = FIXTURE_TOLERANCES["GANGTRAVEL"]
    gang_hoist_tol, _ = FIXTURE_TOLERANCES["GANGHOIST"]
    hoistrow_tol, _ = FIXTURE_TOLERANCES["HOISTROW"]
    bay_travel_tol, _ = FIXTURE_TOLERANCES["BAY_TRAVEL"]
    bay_hoist_tol, _ = FIXTURE_TOLERANCES["LOADBEAM"]
    bay_pos_tol, _ = FIXTURE_TOLERANCES["BAY_POS"]
    th, sh, hh, rh, fhr, fhh, rowgap_cells = _fixture_check_hoist_between_rows(df, max_row)
    tbt, sbt, hbt, rbt, fbtr, fbth, baytravel_cells = _fixture_check_bay_pairs(df, max_row, max_col_map, "TRAVEL")
    tbh, sbh, hbh, rbh, fbhr, fbhh, bayhoist_cells = _fixture_check_bay_pairs(df, max_row, max_col_map, "HOIST")
    tbp, sbp, hbp, rbp, fbpr, fbph, baypos_cells = _fixture_check_bay_pairs(df, max_row, max_col_map, "POS")

    ws = wb.create_sheet("Fixt_XYZ_Difference")
    _title(ws, "RTA / HTA fixture checks — X / Y / Z fork-pair differences", 1, kind="title",
           span=max(len(h_x), len(h_y), len(h_z)))
    r = 3
    r = _write_grouped_fixture_section(ws, r, t_x, s_x, h_x, r_x)
    r = _write_grouped_fixture_section(ws, r, t_y, s_y, h_y, r_y)
    r = _write_grouped_fixture_section(ws, r, t_z, s_z, h_z, r_z)
    _finish_sheet(ws, "A2")

    ws = wb.create_sheet("Fixt_Gang")
    _title(ws, "RTA / HTA fixture checks — gang pick/place  (Level × column-pair, mm error vs spec)",
           1, kind="title", span=8)
    r = 3
    r = _write_heat_matrix_sections(ws, r, "Gang travel — |F1Travel@c1 − F2Travel@c2|", gang_travel_cells, gang_travel_tol)
    r = _write_heat_matrix_sections(ws, r, "Gang hoist — |F1Hoist@c1 − F2Hoist@c2|", gang_hoist_cells, gang_hoist_tol)

    ws = wb.create_sheet("Fixt_Spacing")
    _title(ws, "RTA / HTA fixture checks — spacing  (mm error vs spec; green ≤ spec)", 1, kind="title", span=8)
    r = 3
    r = _write_heat_matrix_sections(ws, r, "Hoist row-gap — |F@row + spacing − F@row+1|  (Level-transition × Column)", rowgap_cells, hoistrow_tol)
    r = _write_heat_matrix_sections(ws, r, "Bay Travel (X) level — all column pairs + bay gap |ΔTravel − nominal|  (Level × column-pair, Fork1/Fork2)", baytravel_cells, bay_travel_tol)
    r = _write_heat_matrix_sections(ws, r, "Bay Hoist (Z) level — all column pairs |FHoist@a − FHoist@b|  (Level × column-pair, Fork1/Fork2)", bayhoist_cells, bay_hoist_tol)
    r = _write_heat_matrix_sections(ws, r, "Bay Pos (Y) level — all column pairs |FPos@a − FPos@b|  (Level × column-pair, Fork1/Fork2)", baypos_cells, bay_pos_tol)

    summary.update(
        {
            "Y_failed_RTA": fy_r, "Y_failed_HTA": fy_h,
            "Z_failed_RTA": fz_r, "Z_failed_HTA": fz_h,
            "X_failed_RTA": fx_r, "X_failed_HTA": fx_h,
            "Gang_travel_failed_RTA": fgrt_r, "Gang_travel_failed_HTA": fgrt_h,
            "Gang_travel_pairs_RTA": tprt, "Gang_travel_pairs_HTA": tpht,
            "Gang_travel_buckets": gtb,
            "Gang_hoist_failed_RTA": fgrh_r, "Gang_hoist_failed_HTA": fgrh_h,
            "Gang_hoist_pairs_RTA": tphr, "Gang_hoist_pairs_HTA": tphh,
            "Gang_hoist_buckets": ghb,
            "Hoist_row_gap_failed_RTA": fhr, "Hoist_row_gap_failed_HTA": fhh,
            "Bay_travel_failed_RTA": fbtr, "Bay_travel_failed_HTA": fbth,
            "Bay_hoist_failed_RTA": fbhr, "Bay_hoist_failed_HTA": fbhh,
            "Bay_pos_failed_RTA": fbpr, "Bay_pos_failed_HTA": fbph,
        }
    )

    _build_summary_sheet(
        wb,
        [
            ("X separation", h_x, r_x),
            ("Y difference", h_y, r_y),
            ("Z difference", h_z, r_z),
            ("Gang travel", hg_t, rg_t),
            ("Gang hoist", hg_h, rg_h),
            ("Hoist row-gap", hh, rh),
            ("Bay Travel level", hbt, rbt),
            ("Bay Hoist level", hbh, rbh),
            ("Bay Pos level", hbp, rbp),
        ],
        int(len(df)),
        max_row,
        aisle,
        even,
    )
    return summary


# ======================================================================================
# Workbook assembly and save
# ======================================================================================


def _write_all_rack_data_sheet(wb: Any, sheet_name: str, rows: list[dict]) -> None:
    # key layout: Location = key[0:2], Row = key[2:4], Column = key[4:6], Line = key[6:8]
    ws = wb.create_sheet(sheet_name[:31])
    headers = [
        "Location",
        "Row",
        "Column",
        "Line",
        "Fork1TravelPos",
        "Fork2TravelPos",
        "Fork1HoistPos",
        "Fork2HoistPos",
        "Fork1Pos",
        "Fork2Pos",
        "Inhibit",
    ]
    _style_header(ws, headers, row=1)
    for i, r in enumerate(rows, start=2):
        _write_row(
            ws,
            i,
            [
                r["location"],
                r["level"],
                r["col"],
                r["side"],
                r["fork1_tp"],
                r["fork2_tp"],
                r["fork1_hp"],
                r["fork2_hp"],
                r["fork1_pp"],
                r["fork2_pp"],
                r["inhibit"],
            ],
        )
    _set_widths(ws, [9, 6, 8, 6, 14, 14, 14, 14, 11, 11, 9])
    ws.freeze_panes = "A2"


def _write_extreme_deltas_sheet(wb: Any, analysis_rows: list[dict]) -> None:
    ws = wb.create_sheet("Extreme_Deltas")
    _title(ws, "Extreme deltas — top 30 per location by |deviation|", 1, kind="title", span=6)
    _title(
        ws,
        "Travel = Fork2Travel − (Fork1Travel + 1180).  Hoist = Fork2Hoist − Fork1Hoist.  "
        "Pos = Fork2Pos − Fork1Pos.  (Current location filter.)",
        2,
        kind="subtitle",
        span=6,
    )
    metrics = (
        ("Travel", "d_travel_err", "Travel_dev_mm", FIXTURE_TOLERANCES["X_DIFF"][0]),
        ("Hoist", "d_hoist", "Hoist_dev_mm", FIXTURE_TOLERANCES["Z_DIFF"][0]),
        ("Pos", "d_pos", "Pos_dev_mm", FIXTURE_TOLERANCES["Y_DIFF"][0]),
    )
    r = 4
    for loc in sorted({x["location"] for x in analysis_rows}):
        rows_loc = [x for x in analysis_rows if x["location"] == loc]
        _title(ws, f"Location {loc} — {LOCATION_NAMES.get(loc, loc)}", r, kind="section", span=6)
        r += 1
        if not rows_loc:
            ws.cell(
                row=r,
                column=1,
                value=_cell_value_for_excel("No rows for this location in the current filter."),
            )
            r += 2
            continue
        for mlabel, key, valhdr, spec in metrics:
            top = sorted(rows_loc, key=lambda x: abs(x[key]), reverse=True)[:30]
            _title(ws, f"{mlabel} — top {len(top)} by |{valhdr}|", r, kind="subtitle", span=6)
            r += 1
            r = _style_header(ws, ["Key", "Location", "Side", "Level", "Col", valhdr], row=r)
            vmin, vmax = _scale_bounds([x[key] for x in top if abs(x[key]) <= spec])
            for x in top:
                _write_row(
                    ws,
                    r,
                    [x["key"], x["location"], x["side"], x["level"], x["col"], round(x[key], 4)],
                )
                v = x[key]
                fill = _scale_fill(v, vmin, vmax, abs(v) > spec)
                if fill is not None:
                    ws.cell(row=r, column=6).fill = fill
                r += 1
            r += 1
        r += 1
    _set_widths(ws, [12, 9, 6, 7, 6, 16])
    _finish_sheet(ws, "A4")


def _build_workbook(
    json_path: str,
    all_rows_raw: list[dict],
    all_rows: list[dict],
    filtered_rows: list[dict],
    analysis_rows: list[dict],
    skipped_fork_unreachable: int,
    location_filter: str,
    skipped_bad_key: int,
    skipped_missing: int,
    aisle: int,
    even: bool,
    offsets: LaserPaddleOffsets | None = None,
) -> tuple[Any, dict[str, Any]]:
    try:
        from openpyxl import Workbook
    except ImportError:
        sys.exit(1)

    wb = Workbook()
    default_ws = wb.active

    _write_all_rack_data_sheet(wb, "All_Rack_Data", all_rows_raw)
    if offsets is not None:
        _write_all_rack_data_sheet(wb, "All_Rack_Data_Offsets", all_rows)

    m_travel = (
        "Each cell = mean of [Fork2Travel_actual − (Fork1Travel + 1180 mm)] in mm "
        "for keys in this block, grouped by Level (row) and Column (col)."
    )
    m_hoist = (
        "Each cell = mean of (Fork2Hoist − Fork1Hoist) in mm, "
        "grouped by Level (row) and Column (col)."
    )
    m_pos = (
        "Each cell = mean of (Fork2Pos − Fork1Pos), grouped by Level (row) and Column (col)."
    )
    travel_sheet_sub = (
        "Pivot-style tables by Location (01–05) and Side: on 01, 03, 04 use side 10 (Non-CD) and 40 (CD). "
        "On aisles 02 and 05 only side 30 (CD) appears here. Empty Location+Side blocks show a short note."
    )
    hoist_sheet_sub = (
        "Same layout as Travel_analysis: stacked Level×Column mean (Fork2Hoist − Fork1Hoist) per Location and Side."
    )
    pos_sheet_sub = (
        "Same layout as Travel_analysis: stacked Level×Column mean (Fork2Pos − Fork1Pos) per Location and Side."
    )

    ws_tr = wb.create_sheet("Travel_analysis")
    _fill_combined_location_side_matrices(
        ws_tr,
        analysis_rows,
        "d_travel_err",
        m_travel,
        "Travel deviation analysis (by Location and Side)",
        travel_sheet_sub,
        FIXTURE_TOLERANCES["X_DIFF"][0],
    )

    ws_ho = wb.create_sheet("Hoist_analysis")
    _fill_combined_location_side_matrices(
        ws_ho,
        analysis_rows,
        "d_hoist",
        m_hoist,
        "Hoist delta analysis (by Location and Side)",
        hoist_sheet_sub,
        FIXTURE_TOLERANCES["Z_DIFF"][0],
    )

    ws_po = wb.create_sheet("Pos_analysis")
    _fill_combined_location_side_matrices(
        ws_po,
        analysis_rows,
        "d_pos",
        m_pos,
        "Pos delta analysis (by Location and Side)",
        pos_sheet_sub,
        FIXTURE_TOLERANCES["Y_DIFF"][0],
    )

    rta_fork_rows = _rta_rows_for_fork_tabs(filtered_rows, all_rows_raw, even)
    _write_fork_raw_tab(
        wb,
        "Fork1_RackData_Analysis",
        "Fork1",
        rta_fork_rows,
        1,
        FORK1_UNREACHABLE_COLS,
        "Columns 01/02 excluded (Fork1 unreachable there).",
    )
    _write_fork_raw_tab(
        wb,
        "Fork2_RackData_Analysis",
        "Fork2",
        rta_fork_rows,
        2,
        FORK2_UNREACHABLE_COLS,
        "Columns 71/72 excluded (Fork2 unreachable there).",
    )

    _write_extreme_deltas_sheet(wb, analysis_rows)

    fix_summary = _append_fixture_sheets_to_workbook(wb, all_rows, even, aisle)

    wb.remove(default_ws)
    return wb, fix_summary


def _save_workbook_with_fallback(wb: Any, target_dir: str, basename: str) -> str:
    base, ext = os.path.splitext(basename)
    candidates = [basename] + [f"{base}_{i}{ext}" for i in range(1, 51)]
    last_err: OSError | None = None
    for name in candidates:
        path = os.path.join(target_dir, name)
        try:
            wb.save(path)
            if name != basename:
                print(
                    f"NOTE: '{basename}' was locked (open in Excel/OneDrive); "
                    f"saved to '{name}' instead."
                )
            return path
        except PermissionError as e:
            last_err = e
            continue
    raise last_err if last_err is not None else OSError("Could not save workbook.")


# ======================================================================================
# Entry point
# ======================================================================================


def main() -> None:
    loc_arg: str | None = None
    aisle_arg: int | None = None
    if len(sys.argv) > 1:
        json_path = os.path.abspath(os.path.expanduser(sys.argv[1].strip('"')))
        if len(sys.argv) > 2:
            loc_arg = sys.argv[2].strip().upper()
        if len(sys.argv) > 3:
            try:
                aisle_arg = int(sys.argv[3].strip())
            except ValueError:
                aisle_arg = None
    else:
        json_path = os.path.abspath(os.path.expanduser(_pick_rack_json_path()))

    if not os.path.isfile(json_path):
        sys.exit(1)

    target_dir = os.path.dirname(json_path)

    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    detected_aisle = _parse_aisle_from_path(json_path)
    if aisle_arg is not None:
        aisle: int | None = aisle_arg
    else:
        aisle = _pick_aisle(detected_aisle)
    if aisle is None:
        sys.exit(1)
    even = aisle % 2 == 0
    print(f"Aisle {aisle} -> CD-side layout: {'EVEN' if even else 'ODD'}")

    if aisle_arg is not None:
        offsets = _load_offsets_for_aisle(aisle)  # non-interactive: apply saved offsets for this aisle
    else:
        offsets = _pick_offsets_for_aisle(aisle)
    _print_offsets_applied(offsets)

    all_rows_raw, skipped_bad, skipped_missing = _parse_rack_rows(data, None)
    all_rows, _, _ = _parse_rack_rows(data, offsets)
    if not all_rows_raw:
        sys.exit(1)

    locations = sorted({r["location"] for r in all_rows})

    loc_sel: LocationSelection | None = None
    if loc_arg is None:
        loc_sel = _pick_location_filter(locations)
        if loc_sel is None:
            sys.exit(1)
    else:
        try:
            loc_sel = _parse_location_cli(loc_arg, set(locations))
        except ValueError:
            sys.exit(1)

    filtered_rows = _filter_by_locations(all_rows, loc_sel)
    location_filter = _format_location_selection(loc_sel)

    if not filtered_rows:
        sys.exit(1)

    analysis_rows, skipped_fork_unreachable = _analysis_rows_for_workbook(filtered_rows, even)
    if not analysis_rows:
        sys.exit(1)

    wb, _fix_summary = _build_workbook(
        json_path,
        all_rows_raw,
        all_rows,
        filtered_rows,
        analysis_rows,
        skipped_fork_unreachable,
        location_filter,
        skipped_bad,
        skipped_missing,
        aisle,
        even,
        offsets,
    )
    xlsx_path = _save_workbook_with_fallback(wb, target_dir, _output_basename(aisle))

    print(f"Excel file created successfully. Saved to: {xlsx_path}")


if __name__ == "__main__":
    main()
