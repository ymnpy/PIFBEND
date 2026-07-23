"""
PBARL Flange / Skin / Web Neighbor Extractor - CLI Edition
============================================================================
Same extraction logic as before, minus the Qt/Win95 GUI. Run it from a
terminal, pass file paths as arguments (or let it prompt you), and watch
a plain-text progress bar do its thing under a chunky ASCII banner.

Optionally, drop a toolbox (.xlsm) on it too: the extracted rows get
pasted into a toolbox sheet, the workbook is recalculated via xlwings,
and the results (including an RF min summary + Info sheet) are written
out - same pattern as the TATOBIN rf_* toolbox functions.

Dependencies beyond the original (pyNastran, pandas, numpy, openpyxl):
    pip install pyfiglet tqdm xlwings

Usage:
    python pbarl_neighbor_extractor_cli.py \
        --bdf model.bdf --load loads.csv --misc misc.xlsx [--out results.xlsx]

    With toolbox integration:
    python pbarl_neighbor_extractor_cli.py \
        --bdf model.bdf --load loads.csv --misc misc.xlsx --toolbox ttb.xlsm

    Or just run it with no arguments and answer the prompts. Files can
    also be dragged-and-dropped onto the script/its .bat launcher.
"""

import argparse
import datetime
import os
import sys
import time

import numpy as np
import pandas as pd
from pyNastran.bdf.bdf import read_bdf

try:
    import pyfiglet
    _HAVE_FIGLET = True
except ImportError:
    _HAVE_FIGLET = False

try:
    from tqdm import tqdm
    _HAVE_TQDM = True
except ImportError:
    _HAVE_TQDM = False

try:
    import xlwings as xw
    _HAVE_XLWINGS = True
except ImportError:
    _HAVE_XLWINGS = False

# ---------------------------------------------------------------------
# CONFIG - adjust to match your actual column names
# ---------------------------------------------------------------------
COL_LOAD_BAR_ELEMENT = "Bar Element ID"
COL_LOAD_BAR_PROPERTY = "Bar Property ID"
COL_LOAD_SUBCASE = "Subcase ID"
COL_LOAD_FX = "FX"
COL_LOAD_FY = "FY"
COL_LOAD_FZ = "FZ"

# Misc Excel is expected to contain (at least) these two sheets:
MISC_SHEET_GENERAL = "GENERAL"
MISC_SHEET_JOINT = "JOINT"

COL_MISC_PBARL = "PBARL"
COL_MISC_BAR_ELEMENT = "PBARL ELEMENT"
COL_MISC_T = "T"
COL_MISC_W = "W"
COL_MISC_MAT = "MAT"
COL_MISC_PSHELL = "PSHELL"
COL_MISC_T_WEB = "T WEB"
COL_MISC_PLIES = ["N1", "N2", "N3", "N4"]
COL_MISC_X_LENGTH = "X-length"
COL_MISC_Y_LENGTH = "Y-length"

SKIN_MARKER_COL = "NAME"
SKIN_MARKER_VALUE = "M91"

# JOINT sheet - one row per bar element with fastener/joint properties.
# Assumed to share the same element-id column name as the load CSV;
# edit COL_JOINT_ELEMENT if the JOINT sheet uses a different header.
COL_JOINT_ELEMENT = "Bar Element ID"
COL_JOINT_DIAMETER = "diameter"
COL_JOINT_PITCH = "pitch"
COL_JOINT_EX = "ex"

ANGLE_TOL_DEG = 45.0
MIN_SHARED_NODES = 2
MAX_HOPS = 2  # bar -> first shell -> second shell, then stop (don't keep wandering)

# Property types that are actually shells, i.e. valid candidates for
# material_coordinate_system(). Anything else (bar/beam/rod/etc. props
# that can slip through the shared-node neighbor search) gets routed
# around instead of crashing check_angle().
SHELL_PROPERTY_TYPES = {"PSHELL", "PCOMP", "PCOMPG"}


# ---------------------------------------------------------------------
# TTB (TOOLBOX) CONFIG - *** EDIT THESE TO MATCH YOUR ACTUAL TOOLBOX ***
# ---------------------------------------------------------------------
# Name of the sheet/tab in the toolbox workbook that this analysis writes
# to and reads results from (same idea as the ttb_sheet variable in each
# rf_* function of the TATOBIN example).
TTB_SHEET_NAME = "PBARL Neighbor - Automation"  # TODO: put your real tab name here

# First data row when there's at least one row to write, and the row used
# instead when the input is empty (mirrors the `start = X if rows!=0 else Y`
# pattern used throughout the example - some toolbox sheets have an extra
# header/example row that needs to be skipped when there's no real data).
TTB_START_ROW = 10          # TODO
TTB_START_ROW_EMPTY = 12    # TODO

# Full column range that gets AutoFill'd/cleared on the toolbox sheet -
# should span from the first input column through the last output/RF
# column, e.g. ("B", "AZ").
TTB_FIRST_COL = "B"   # TODO
TTB_LAST_COL = "AZ"   # TODO

# Maps a column from the extractor's df_out to the toolbox INPUT column
# letter it should be written to. Add/remove rows to match your sheet.
# Left side = column name produced by run_extraction(); right side =
# spreadsheet column letter on TTB_SHEET_NAME.
TTB_INPUT_COLUMNS = {
    # df_out column          : toolbox column
    "MAT_SKIN":               "B",   # TODO
    "N1":                     "C",   # TODO
    "N2":                     "D",   # TODO
    "N3":                     "E",   # TODO
    "N4":                     "F",   # TODO
    "X_LENGTH":                "G",   # TODO
    "Y_LENGTH":                "H",   # TODO
    "T_WEB":                  "I",   # TODO
    "MAT_WEB":                "J",   # TODO
    "DIAMETER":               "K",   # TODO
    "PITCH":                  "L",   # TODO
    "EX":                     "M",   # TODO
    "FX":                     "N",   # TODO
    "FY":                     "O",   # TODO
    "FZ":                     "P",   # TODO
}

# Column in the toolbox output range that holds the resulting reserve
# factor - used to build the "RF min" summary sheet. Set to None to skip
# the RF-min summary if this analysis doesn't produce one.
TTB_RF_COLUMN = "RF"  # TODO (must match a header the toolbox writes out)


# ---------------------------------------------------------------------
# CORE LOGIC (unchanged from the GUI version)
# ---------------------------------------------------------------------
def unit(v):
    v = np.asarray(v, dtype=float)
    n = np.linalg.norm(v)
    return v / n if n > 0 else v


def build_property_node_map(bdf):
    prop_nodes = {}
    for eid, elem in bdf.elements.items():
        pid = getattr(elem, "pid", None)
        if pid is None or pid not in bdf.properties:
            continue
        prop_nodes.setdefault(pid, set()).update(elem.node_ids)
    return prop_nodes


def build_property_type_map(bdf):
    return {pid: prop.type for pid, prop in bdf.properties.items()}


def find_neighbors(pid, prop_nodes, exclude=None, min_shared=MIN_SHARED_NODES):
    exclude = exclude or set()
    if pid not in prop_nodes:
        return []
    nodes = prop_nodes[pid]
    return [
        other_pid
        for other_pid, other_nodes in prop_nodes.items()
        if other_pid != pid
        and other_pid not in exclude
        and len(nodes & other_nodes) >= min_shared
    ]


def build_property_element_count(bdf):
    """{pid: number of elements referencing that property}."""
    counts = {}
    for elem in bdf.elements.values():
        pid = getattr(elem, "pid", None)
        if pid is None:
            continue
        counts[pid] = counts.get(pid, 0) + 1
    return counts


def find_property_neighbors_by_type(pid, prop_nodes, prop_types, target_types,
                                     exclude=None, min_shared=2):
    """
    Neighbors of `pid` (by shared nodes, >=min_shared), restricted to
    properties whose type is in `target_types`.
    """
    exclude = exclude or set()
    if pid not in prop_nodes:
        return []
    nodes = prop_nodes[pid]
    return [
        other_pid
        for other_pid, other_nodes in prop_nodes.items()
        if other_pid != pid
        and other_pid not in exclude
        and prop_types.get(other_pid) in target_types
        and len(nodes & other_nodes) >= min_shared
    ]


def find_deepest_shell(start_pid, prop_nodes, prop_types,
                        exclude_types=None, max_hops=MAX_HOPS):
    """
    Walk from start_pid inward through shared-node neighbors, one hop at
    a time, never revisiting a property already seen.

    At each hop:
      1. Try an unvisited neighbor sharing >=2 nodes (normal edge-shared).
      2. If none, fall back to a neighbor sharing exactly 1 node
         (corner-only touch).
      3. Stop when nothing further is found, or after max_hops steps.

    Returns (final_pid, path).
    """
    exclude_types = exclude_types or set()
    visited = [start_pid]
    current = start_pid

    for _ in range(max_hops):
        nodes = prop_nodes.get(current, set())

        def _search(threshold):
            for other_pid, other_nodes in prop_nodes.items():
                if other_pid in visited:
                    continue
                if prop_types.get(other_pid) in exclude_types:
                    continue
                if len(nodes & other_nodes) >= threshold:
                    return other_pid
            return None

        next_pid = _search(2)
        if next_pid is None:
            next_pid = _search(1)
        if next_pid is None:
            break

        visited.append(next_pid)
        current = next_pid

    return current, visited


def resolve_pbarl_to_shell(pbarl_pid, prop_nodes, prop_types, prop_elem_count):
    """
    Step A: if pbarl_pid itself has only 1 element (a corner bar), look
    for a NEIGHBORING PBARL property (sharing >=2 nodes, falling back to
    a corner-only 1-node touch) that has more than 1 element - i.e. a
    "real" edge bar rather than a lone corner stub. If found, reroute to
    that bar as the effective starting point. If pbarl_pid already has
    >1 element, or no such neighbor bar exists, just use pbarl_pid itself.

    Step B: from the effective bar, run find_deepest_shell to reach the
    innermost shell (e.g. the blue interior property in a skin/doubler
    stack), excluding PBARL properties from the walk.

    Returns (lookup_pid, effective_bar_pid, hop_path).
    """
    effective_bar = pbarl_pid
    own_count = prop_elem_count.get(pbarl_pid, 0)

    if own_count <= 1:
        # try a normal edge-shared bar neighbor first (>=2 nodes), then
        # fall back to a corner-only touch (1 shared node) - two CBARs
        # meeting end-to-end at a single grid point only share 1 node.
        bar_neighbors = find_property_neighbors_by_type(
            pbarl_pid, prop_nodes, prop_types, target_types={"PBARL"}, min_shared=2
        )
        if not bar_neighbors:
            bar_neighbors = find_property_neighbors_by_type(
                pbarl_pid, prop_nodes, prop_types, target_types={"PBARL"}, min_shared=1
            )

        for candidate in bar_neighbors:
            if prop_elem_count.get(candidate, 0) > 1:
                effective_bar = candidate
                break
        # if no multi-element neighbor bar found, effective_bar stays pbarl_pid

    lookup_pid, hop_path = find_deepest_shell(
        effective_bar, prop_nodes, prop_types, exclude_types={"PBARL"}
    )
    return lookup_pid, effective_bar, hop_path


def get_elements_for_property(bdf, pid):
    return [e for e in bdf.elements.values() if getattr(e, "pid", None) == pid]


def get_shell_x_axis_fallback(element):
    """
    Fallback for when material_coordinate_system() isn't available, raises,
    or returns None (non-shell element, unresolved coord system/nodes,
    degenerate geometry, etc). Approximates the in-plane x-axis directly
    from the element's own node positions instead of relying on pyNastran's
    material frame. Not identical to the real material x-axis (it ignores
    any THETA/MCID rotation on the shell), but good enough to judge whether
    the bar runs roughly parallel or reversed relative to the shell edge.
    """
    nodes_ref = getattr(element, "nodes_ref", None)
    if nodes_ref is None or any(n is None for n in nodes_ref):
        return None
    if len(nodes_ref) < 2:
        return None

    try:
        p0 = nodes_ref[0].get_position()
        p1 = nodes_ref[1].get_position()
    except Exception:
        return None

    x_axis = unit(p1 - p0)
    if np.linalg.norm(x_axis) == 0:
        return None
    return x_axis


def check_angle(bar_element, shell_pid, bdf, tol=ANGLE_TOL_DEG):
    bar_node1 = bar_element.ga_ref.get_position()
    bar_node2 = bar_element.gb_ref.get_position()
    bar_dir = unit(bar_node2 - bar_node1)

    shell_elements = get_elements_for_property(bdf, shell_pid)
    if not shell_elements:
        return False

    element = shell_elements[0]
    x_axis = None

    # Only true shell elements implement material_coordinate_system(). Even
    # when they do, it can still come back None for degenerate geometry or
    # an unresolved coordinate system/node reference, so we guard both the
    # attribute check and the call itself.
    if hasattr(element, "material_coordinate_system"):
        try:
            result = element.material_coordinate_system()
        except Exception as e:
            print(f"[check_angle] material_coordinate_system() failed for "
                  f"element {getattr(element, 'eid', '?')} "
                  f"(pid {shell_pid}): {e}")
            result = None

        if result is not None:
            _theta, _centroid, imat, _jmat, _normal = result
            x_axis = unit(imat)

    if x_axis is None:
        x_axis = get_shell_x_axis_fallback(element)

    if x_axis is None:
        print(f"[check_angle] could not determine orientation for element "
              f"{getattr(element, 'eid', '?')} (pid {shell_pid}) - "
              f"defaulting reverse=False")
        return False

    cos_a = np.clip(np.dot(bar_dir, x_axis), -1.0, 1.0)
    angle_deg = np.degrees(np.arccos(cos_a))
    return angle_deg > tol


def clear_range_dynamic(ws, start_row, start_col, end_col):
    """
    Clear a toolbox range from start_row down to the last used row in
    start_col, so leftover data from a previous run doesn't bleed into
    the next one. No-op if the sheet is already empty there.

    Uses >= rather than > on the boundary check: the original TATOBIN
    helper this is based on used `if last_row > start_row`, which silently
    skips clearing when there's exactly one autofilled row below the
    template row (last_row == start_row) - a real edge case whenever a
    batch produces only 2 total toolbox rows.
    """
    last_row = ws.range(start_col + str(ws.cells.last_cell.row)).end('up').row
    if last_row >= start_row:
        ws.range(f"{start_col}{start_row}:{end_col}{last_row}").clear_contents()


def rf_pbarl_neighbor_extractor(wb, df_out):
    """
    Pushes the extracted PBARL neighbor rows into the toolbox, recalculates,
    and reads the results back - same shape as the rf_* functions in the
    TATOBIN toolbox interface (e.g. rf_local_buckling): clear the target
    range, AutoFill formulas down to the row count, write each input
    column, wb.app.calculate(), then read the whole range back into a
    DataFrame.

    `df_out` is the DataFrame returned by run_extraction(). Column names
    written to the toolbox are controlled by TTB_INPUT_COLUMNS above -
    edit that dict (and TTB_SHEET_NAME/TTB_FIRST_COL/TTB_LAST_COL) to
    match your actual toolbox layout.

    Returns (ttb_sheet_name, df_out_final, elapsed_seconds) - the same
    3-tuple shape every rf_* function in the example returns, so it can
    be handed straight to write_ttb_results().
    """
    start = TTB_START_ROW if df_out.shape[0] != 0 else TTB_START_ROW_EMPTY
    begin = time.time()

    ttb_sheet = TTB_SHEET_NAME
    ws = wb.sheets(ttb_sheet)

    rows = df_out.shape[0]
    rows = rows + start - 1

    clear_range_dynamic(ws, start + 1, TTB_FIRST_COL, TTB_LAST_COL)
    ws.range(f'{TTB_FIRST_COL}{start}:{TTB_LAST_COL}{start}').api.AutoFill(
        Destination=ws.range(f'{TTB_FIRST_COL}{start}:{TTB_LAST_COL}{rows}').api, Type=0
    )

    # --- write each mapped input column ---
    for df_col, ttb_col in TTB_INPUT_COLUMNS.items():
        if df_col not in df_out.columns:
            print(f"[rf_pbarl_neighbor_extractor] WARNING: column '{df_col}' "
                  f"not found in extraction output, skipping.")
            continue
        ws.range(f'{ttb_col}{start}').value = np.array(df_out[df_col]).reshape(-1, 1)

    wb.app.calculate()
    end = time.time()

    columns = ws.range(f'{TTB_FIRST_COL}{start-1}:{TTB_LAST_COL}{start-1}').value
    df_result = pd.DataFrame(
        ws.range(f'{TTB_FIRST_COL}{start}:{TTB_LAST_COL}{rows}').value, columns=columns
    )

    # tack the original identifying columns back on the front, same as
    # the df_addition + pd.concat pattern used throughout the example
    id_cols = [c for c in ("Bar Element ID", "Bar Property ID", "Subcase ID") if c in df_out.columns]
    df_addition = df_out[id_cols].reset_index(drop=True)
    df_out_final = pd.concat((df_addition, df_result), axis=1)

    clear_range_dynamic(ws, start + 1, TTB_FIRST_COL, TTB_LAST_COL)

    dt = end - begin
    return ttb_sheet, df_out_final, dt


def write_ttb_results(sheet_name, df_out_final, dt, toolbox_path, out_dir):
    """
    Writes the toolbox results to a timestamped .xlsx with three sheets:
      - the raw results (sheet named after the toolbox tab)
      - "RF min" - one row per Bar Property ID, the minimum RF found
      - "Info" - a small metadata sheet (elapsed time, timestamp, toolbox file)
    Mirrors write_stuff() from the TATOBIN example. Returns the output path.
    """
    sheet_name_clean = sheet_name.replace(" ", "")[:31]  # Excel sheet name limit
    stamp = datetime.datetime.now().strftime("%H%M%S-%d%m%Y")
    output_file = os.path.join(out_dir, f"PBARL_{sheet_name_clean}_{stamp}.xlsx")

    metadata = pd.DataFrame({
        "Parameter": ["Delta Time (seconds)", "Processing Date", "Toolbox File", "User"],
        "Value": [
            f"{dt:0.2f}",
            datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            os.path.basename(toolbox_path),
            os.environ.get("USERNAME") or os.environ.get("USER") or "unknown",
        ],
    })

    df_rf_min = pd.DataFrame()
    if TTB_RF_COLUMN and TTB_RF_COLUMN in df_out_final.columns and "Bar Property ID" in df_out_final.columns:
        df_out_final[TTB_RF_COLUMN] = pd.to_numeric(df_out_final[TTB_RF_COLUMN], errors="coerce")
        try:
            df_rf_min = df_out_final.loc[
                df_out_final.groupby("Bar Property ID")[TTB_RF_COLUMN].idxmin()
            ].reset_index(drop=True)
        except Exception as e:
            print(YELLOW + f"WARNING: could not build RF min summary: {e}" + RESET)
    elif TTB_RF_COLUMN:
        print(YELLOW + f"WARNING: RF column '{TTB_RF_COLUMN}' not found - "
                        f"skipping RF min sheet (check TTB_RF_COLUMN / TTB_INPUT_COLUMNS)." + RESET)

    with pd.ExcelWriter(output_file) as writer:
        df_out_final.to_excel(writer, sheet_name=sheet_name_clean, index=False)
        if not df_rf_min.empty:
            df_rf_min.to_excel(writer, sheet_name="RF min", index=False)
        metadata.to_excel(writer, sheet_name="Info", index=False)

    return output_file


class ExtractionCancelled(Exception):
    """
    Marker exception, not a real error. Raised inside run_extraction() when
    the user hits Ctrl+C mid-loop, so we can unwind out of the PBARL loop
    immediately instead of grinding through the rest of the list. It's
    caught separately from real errors (bad file, missing column, etc.) so
    a user-initiated stop shows "Cancelled" instead of an error dialog.
    """
    pass


def run_extraction(bdf, df_load, df_misc_general, df_misc_joint, should_continue, progress_cb):
    prop_nodes = build_property_node_map(bdf)
    prop_types = build_property_type_map(bdf)
    prop_elem_count = build_property_element_count(bdf)

    total = len(df_load)
    rows = []

    # Cache the resolved shell per Bar Property ID so repeat subcases for
    # the same bar don't re-run the reroute+walk every row.
    resolved_cache = {}

    for i, (_, load_row) in enumerate(df_load.iterrows(), 1):
        if not should_continue():
            raise ExtractionCancelled()

        bar_eid = load_row[COL_LOAD_BAR_ELEMENT]
        bar_pid = load_row[COL_LOAD_BAR_PROPERTY]
        subcase = load_row[COL_LOAD_SUBCASE]
        fx = load_row[COL_LOAD_FX]
        fy = load_row[COL_LOAD_FY]
        fz = load_row[COL_LOAD_FZ]

        bar_element = bdf.elements.get(bar_eid)
        if bar_element is None:
            progress_cb(i, total)
            continue

        # --- resolve Bar Property ID -> inner skin shell (reroute + walk) ---
        if bar_pid in resolved_cache:
            lookup_pid = resolved_cache[bar_pid]
        else:
            resolved_pid, _effective_bar, _hop_path = resolve_pbarl_to_shell(
                bar_pid, prop_nodes, prop_types, prop_elem_count
            )
            lookup_pid = (
                resolved_pid
                if resolved_pid is not None
                and prop_types.get(resolved_pid) in SHELL_PROPERTY_TYPES
                else None
            )
            resolved_cache[bar_pid] = lookup_pid

        # --- skin shell's N1-N4 / MAT / X-length / Y-length, from GENERAL,
        # keyed by PSHELL ---
        skin_rows = df_misc_general[df_misc_general[COL_MISC_PSHELL] == lookup_pid]
        if not skin_rows.empty:
            skin_row = skin_rows.iloc[0]
            n1, n2, n3, n4 = (skin_row[c] for c in COL_MISC_PLIES)
            skin_mat = skin_row[COL_MISC_MAT]
            x_length = skin_row[COL_MISC_X_LENGTH]
            y_length = skin_row[COL_MISC_Y_LENGTH]
        else:
            n1 = n2 = n3 = n4 = "n/a"
            skin_mat = "n/a"
            x_length = y_length = "n/a"

        # If the bar runs against the shell's material x-axis (rather than
        # parallel to it), X-length and Y-length are swapped to match the
        # bar's actual orientation relative to the shell.
        if lookup_pid is not None and x_length != "n/a" and y_length != "n/a":
            reverse = check_angle(bar_element, lookup_pid, bdf, tol=ANGLE_TOL_DEG)
            if reverse:
                x_length, y_length = y_length, x_length

        # --- web neighbor's thickness / MAT, from GENERAL, keyed by the
        # bar's own PBARL group (not the resolved skin) ---
        bar_group = df_misc_general[df_misc_general[COL_MISC_PBARL] == bar_pid]
        is_skin = (
            bar_group[SKIN_MARKER_COL].astype(str).str.contains(SKIN_MARKER_VALUE, na=False)
        )
        df_web = bar_group[~is_skin]
        if not df_web.empty:
            web_row = df_web.iloc[0]
            web_t = web_row[COL_MISC_T_WEB]
            web_mat = web_row[COL_MISC_MAT]
        else:
            web_t = "n/a"
            web_mat = "n/a"

        # --- joint properties, from JOINT sheet, keyed by element id ---
        joint_rows = df_misc_joint[df_misc_joint[COL_JOINT_ELEMENT] == bar_eid]
        if not joint_rows.empty:
            joint_row = joint_rows.iloc[0]
            diameter = joint_row[COL_JOINT_DIAMETER]
            pitch = joint_row[COL_JOINT_PITCH]
            ex = joint_row[COL_JOINT_EX]
        else:
            diameter = pitch = ex = "n/a"

        rows.append((
            bar_eid, bar_pid, subcase,
            lookup_pid, n1, n2, n3, n4, skin_mat,
            x_length, y_length,
            web_t, web_mat,
            diameter, pitch, ex,
            fx, fy, fz,
        ))

        progress_cb(i, total)

    return pd.DataFrame(rows, columns=[
        "Bar Element ID", "Bar Property ID", "Subcase ID",
        "PSHELL", "N1", "N2", "N3", "N4", "MAT_SKIN",
        "X_LENGTH", "Y_LENGTH",
        "T_WEB", "MAT_WEB",
        "DIAMETER", "PITCH", "EX",
        "FX", "FY", "FZ",
    ])


# ---------------------------------------------------------------------
# TERMINAL THEME - figlet banner + a couple of ANSI helpers
# ---------------------------------------------------------------------
CYAN = "\033[96m"
GREEN = "\033[92m"
YELLOW = "\033[93m"
RED = "\033[91m"
BOLD = "\033[1m"
DIM = "\033[2m"
RESET = "\033[0m"


def print_banner():
    title = "PBARL EXTRACTOR"
    if _HAVE_FIGLET:
        try:
            print(CYAN + pyfiglet.figlet_format(title, font="standard") + RESET)
        except Exception:
            print(CYAN + BOLD + f"== {title} ==" + RESET)
    else:
        print(CYAN + BOLD + f"== {title} ==" + RESET)
        print(DIM + "(install 'pyfiglet' for the full ASCII banner: pip install pyfiglet)" + RESET)
    print(DIM + "Flange / Skin / Web Neighbor Extractor - CLI Edition" + RESET)
    print(DIM + "-" * 68 + RESET)


def prompt_for_path(label, must_exist=True):
    while True:
        path = input(f"{YELLOW}{label}:{RESET} ").strip().strip('"')
        if not path:
            print(RED + "  Path cannot be empty." + RESET)
            continue
        if must_exist and not os.path.exists(path):
            print(RED + f"  File not found: {path}" + RESET)
            continue
        return path


def make_progress_bar(total):
    """
    Returns a callback(done, total) that updates a terminal progress bar.
    Uses tqdm if available, otherwise falls back to a hand-rolled bar.
    """
    if _HAVE_TQDM:
        bar = tqdm(total=total, desc="Extracting", unit="row",
                   bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}]")

        def cb(done, _total):
            bar.n = done
            bar.refresh()
            if done >= _total:
                bar.close()
        return cb

    bar_width = 40

    def cb(done, _total):
        frac = done / _total if _total else 1.0
        filled = int(bar_width * frac)
        bar_str = GREEN + "#" * filled + RESET + "-" * (bar_width - filled)
        pct = int(frac * 100)
        end = "\n" if done >= _total else ""
        print(f"\r  [{bar_str}] {pct:3d}%  ({done}/{_total})", end=end, flush=True)
    return cb


BDF_EXTS = {".bdf", ".dat", ".nas", ".nastran", ".pch"}
LOAD_EXTS = {".csv", ".tsv", ".txt"}
MISC_EXTS = {".xlsx", ".xls"}
TOOLBOX_EXTS = {".xlsm"}  # macro-enabled workbook - kept distinct from misc so drag-and-drop can tell them apart


def _clean_path(p):
    """Strip the surrounding quotes/whitespace Windows drag-and-drop adds."""
    return p.strip().strip('"').strip("'")


def parse_args():
    parser = argparse.ArgumentParser(
        description="Extract PBARL flange/skin/web neighbor data from a BDF + load CSV + misc Excel. "
                     "You can also just drag and drop the three files onto this script/its .bat launcher."
    )
    parser.add_argument("--bdf", help="Path to the BDF file")
    parser.add_argument("--load", help="Path to the Load CSV file")
    parser.add_argument("--misc", help="Path to the Misc Excel file")
    parser.add_argument("--toolbox", default=None,
                         help="Optional: path to a .xlsm toolbox to push results into via xlwings "
                              "(see rf_pbarl_neighbor_extractor / TTB_* config at the top of this file)")
    parser.add_argument("--out", default=None,
                         help="Output .xlsx path (default: ./pbarl_neighbor_results.xlsx)")
    # Positional catch-all: this is what actually receives the paths when
    # files are dragged-and-dropped onto the script/launcher in Windows
    # Explorer - it hands them over as bare arguments, not --flags.
    parser.add_argument("dropped_files", nargs="*",
                         help=argparse.SUPPRESS)
    return parser.parse_args()


def sort_dropped_files(paths):
    """
    Given a list of raw paths (as Windows drag-and-drop hands them over),
    sort them into (bdf, load, misc) by file extension. Returns a dict
    with any slot left as None if no matching file was dropped.
    """
    result = {"bdf": None, "load": None, "misc": None, "toolbox": None}
    leftovers = []
    for raw in paths:
        p = _clean_path(raw)
        ext = os.path.splitext(p)[1].lower()
        if ext in BDF_EXTS and result["bdf"] is None:
            result["bdf"] = p
        elif ext in TOOLBOX_EXTS and result["toolbox"] is None:
            result["toolbox"] = p
        elif ext in MISC_EXTS and result["misc"] is None:
            result["misc"] = p
        elif ext in LOAD_EXTS and result["load"] is None:
            result["load"] = p
        else:
            leftovers.append(p)
    if leftovers:
        print(YELLOW + f"Ignored unrecognized file(s): {', '.join(leftovers)}" + RESET)
    return result


def main():
    print_banner()
    args = parse_args()

    dropped = sort_dropped_files(args.dropped_files) if args.dropped_files else {}

    bdf_path = (_clean_path(args.bdf) if args.bdf else None) or dropped.get("bdf") \
        or prompt_for_path("BDF file path")
    load_path = (_clean_path(args.load) if args.load else None) or dropped.get("load") \
        or prompt_for_path("Load CSV path")
    misc_path = (_clean_path(args.misc) if args.misc else None) or dropped.get("misc") \
        or prompt_for_path("Misc Excel path")
    toolbox_path = (_clean_path(args.toolbox) if args.toolbox else None) or dropped.get("toolbox")
    out_path = (_clean_path(args.out) if args.out else None) \
        or os.path.join(os.getcwd(), "pbarl_neighbor_results.xlsx")

    required = (("BDF", bdf_path), ("Load CSV", load_path), ("Misc Excel", misc_path))
    for label, path in required:
        if not os.path.exists(path):
            print(RED + f"{label} file not found: {path}" + RESET)
            sys.exit(1)
    if toolbox_path and not os.path.exists(toolbox_path):
        print(RED + f"Toolbox file not found: {toolbox_path}" + RESET)
        sys.exit(1)

    try:
        print(DIM + "Reading BDF..." + RESET)
        bdf = read_bdf(bdf_path, debug=False)

        print(DIM + "Reading Load CSV..." + RESET)
        df_load = pd.read_csv(load_path)

        print(DIM + "Reading Misc Excel (GENERAL, JOINT sheets)..." + RESET)
        misc_sheets = pd.read_excel(misc_path, sheet_name=None)
        df_misc_general = misc_sheets[MISC_SHEET_GENERAL]
        df_misc_joint = misc_sheets[MISC_SHEET_JOINT]

        print()
        progress_cb = make_progress_bar(len(df_load))

        df_out = run_extraction(
            bdf, df_load, df_misc_general, df_misc_joint,
            should_continue=lambda: True,  # Ctrl+C handled via KeyboardInterrupt below
            progress_cb=progress_cb,
        )

        df_out.to_excel(out_path, index=False)
        print()
        print(GREEN + BOLD + f"Done - {len(df_out)} rows -> {out_path}" + RESET)

        # --- optional: push results into a toolbox and recalculate ---
        if toolbox_path:
            if not _HAVE_XLWINGS:
                print(RED + "Toolbox given but 'xlwings' isn't installed. "
                             "Run: pip install xlwings" + RESET)
                sys.exit(1)

            print()
            print(CYAN + f"Pushing results into toolbox: {toolbox_path}" + RESET)
            wb = xw.Book(toolbox_path)
            try:
                try:
                    wb.app.api.Calculation = xw.constants.Calculation.xlCalculationManual
                    wb.app.screen_updating = False
                    wb.app.display_alerts = False
                except Exception:
                    pass

                sheet_name, df_out_final, dt = rf_pbarl_neighbor_extractor(wb, df_out)
                out_dir = os.path.dirname(out_path) or os.getcwd()
                ttb_out_path = write_ttb_results(sheet_name, df_out_final, dt, toolbox_path, out_dir)

                print(GREEN + BOLD + f"Toolbox done ({dt:0.2f}s) -> {ttb_out_path}" + RESET)
            finally:
                try:
                    wb.app.api.Calculation = xw.constants.Calculation.xlCalculationAutomatic
                    wb.app.screen_updating = True
                    wb.app.display_alerts = True
                except Exception:
                    pass

    except KeyboardInterrupt:
        print()
        print(YELLOW + "Cancelled by user." + RESET)
        sys.exit(130)
    except ExtractionCancelled:
        print()
        print(YELLOW + "Cancelled." + RESET)
        sys.exit(130)
    except KeyError as e:
        print()
        print(RED + f"Missing expected sheet or column: {e}" + RESET)
        sys.exit(1)
    except Exception as e:
        print()
        print(RED + f"Error: {e}" + RESET)
        sys.exit(1)


if __name__ == "__main__":
    main()
    # If launched by dropping files onto the script/.bat (rather than from
    # an already-open terminal), Windows closes the console the instant the
    # script exits - so pause here to keep the result visible. Guarded
    # against EOFError so running this non-interactively (piped stdin,
    # scheduled tasks, CI, etc.) doesn't turn a successful run into a crash.
    if len(sys.argv) > 1:
        try:
            input(DIM + "\nPress Enter to close..." + RESET)
        except (EOFError, KeyboardInterrupt):
            pass
