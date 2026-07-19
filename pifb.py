"""
PBARL Flange / Skin / Web Neighbor Extractor - basic GUI
=========================================================
Plain tkinter, no styling library. Three file inputs, a progress bar,
Run / Cancel, and a small image slot on the left. Column-name assumptions
are in the CONFIG block below - edit those if your CSV/Excel headers differ.
"""

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import numpy as np
import pandas as pd
from pyNastran.bdf.bdf import read_bdf

try:
    from PIL import Image, ImageTk, ImageOps
    HAS_PIL = True
except ImportError:
    HAS_PIL = False


def _load_logo_image(path, max_size):
    """
    Load an image and make it exactly fill max_size=(w, h) - no empty
    letterbox space, no distortion. Achieved by scaling to cover the box
    and center-cropping any excess, same idea as CSS 'object-fit: cover'.
    Returns a Tk-compatible image object.
    """
    if HAS_PIL:
        img = Image.open(path)
        # Photos from phones/cameras often carry an EXIF rotation tag rather
        # than being physically rotated - apply it so the preview matches
        # what you'd see in a normal image viewer.
        img = ImageOps.exif_transpose(img)
        # ImageOps.fit scales the image up/down so it fully covers max_size,
        # then center-crops whatever sticks out past the target box.
        img = ImageOps.fit(img, max_size, Image.LANCZOS)
        return ImageTk.PhotoImage(img)

    # No Pillow: tk.PhotoImage has no crop/cover operation, only whole-number
    # subsample() downscaling - so this can only approximate "fill" by
    # shrinking to just above the target size, not an exact pixel match.
    img = tk.PhotoImage(file=path)
    if img.width() > max_size[0] or img.height() > max_size[1]:
        factor = max(
            img.width() // max_size[0] if max_size[0] else 1,
            img.height() // max_size[1] if max_size[1] else 1,
            1,
        )
        img = img.subsample(factor, factor)
    return img

# ---------------------------------------------------------------------
# CONFIG - adjust to match your actual column names
# ---------------------------------------------------------------------
COL_LOAD_PID = "PID"
COL_LOAD_A = "a"
COL_LOAD_B = "b"

COL_MISC_PBARL = "PBARL"
COL_MISC_BAR_ELEMENT = "PBARL ELEMENT"
COL_MISC_T = "T"
COL_MISC_W = "W"
COL_MISC_MAT = "MAT"
COL_MISC_PSHELL = "PSHELL"
COL_MISC_T_WEB = "T WEB"
COL_MISC_PLIES = ["N1", "N2", "N3", "N4"]

SKIN_MARKER_COL = "NAME"
SKIN_MARKER_VALUE = "M91"

ANGLE_TOL_DEG = 45.0
MIN_SHARED_NODES = 2

# Point this at your own image. With Pillow installed, any common format
# (JPEG, PNG, BMP, ...) works and gets scaled to fit IMAGE_PANEL_SIZE below
# while keeping its aspect ratio. Without Pillow, only PNG/GIF work and
# scaling is limited to integer downscaling (see _load_logo_image below).
# Leave IMAGE_PATH as "" to skip it entirely.
IMAGE_PATH = ''
IMAGE_PANEL_SIZE = (100, 240)  # (width, height) the image should fit inside


# ---------------------------------------------------------------------
# CORE LOGIC
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


def get_elements_for_property(bdf, pid):
    return [e for e in bdf.elements.values() if getattr(e, "pid", None) == pid]


def check_angle(bar_element, shell_pid, bdf, tol=ANGLE_TOL_DEG):
    bar_node1 = bar_element.ga_ref.get_position()
    bar_node2 = bar_element.gb_ref.get_position()
    bar_dir = unit(bar_node2 - bar_node1)

    shell_elements = get_elements_for_property(bdf, shell_pid)
    if not shell_elements:
        return False

    element = shell_elements[0]
    theta, centroid, imat, jmat, normal = element.material_coordinate_system()
    x_axis = unit(imat)

    cos_a = np.clip(np.dot(bar_dir, x_axis), -1.0, 1.0)
    angle_deg = np.degrees(np.arccos(cos_a))
    return angle_deg > tol


class ExtractionCancelled(Exception):
    """
    Marker exception, not a real error. Raised inside run_extraction() when
    the user hits Cancel mid-loop, so we can unwind out of the PBARL loop
    immediately instead of grinding through the rest of the list. It's
    caught separately from real errors (bad file, missing column, etc.) so
    a user-initiated stop shows "Cancelled" instead of an error dialog.
    """
    pass


def run_extraction(bdf, df_load, df_misc, should_continue, progress_cb):
    prop_nodes = build_property_node_map(bdf)
    prop_types = build_property_type_map(bdf)

    pbarl_ids = list(pd.unique(df_load[COL_LOAD_PID]))
    total = len(pbarl_ids)
    rows = []

    for i, pbarl in enumerate(pbarl_ids, 1):
        if not should_continue():
            raise ExtractionCancelled()

        df_filter = df_misc[df_misc[COL_MISC_PBARL] == pbarl]
        if df_filter.empty:
            progress_cb(i, total)
            continue

        bar_eid = df_filter[COL_MISC_BAR_ELEMENT].iloc[0]
        bar_element = bdf.elements.get(bar_eid)
        if bar_element is None:
            progress_cb(i, total)
            continue

        T = df_filter[COL_MISC_T].max()
        W = df_filter[COL_MISC_W].max()
        mat = df_filter[COL_MISC_MAT].iloc[0]

        is_skin = (
            df_filter[SKIN_MARKER_COL]
            .astype(str)
            .str.contains(SKIN_MARKER_VALUE, na=False)
        )
        df_skin = df_filter[is_skin]
        df_web = df_filter[~is_skin]

        if df_web.empty:
            web_id, web_t = "n/a", "n/a"
        else:
            web_id = df_web[COL_MISC_PSHELL].iloc[0]
            web_t = df_web[COL_MISC_T_WEB].max()

        for _, skin_row in df_skin.iterrows():
            neighbor_org = skin_row[COL_MISC_PSHELL]
            neighbor_plies = [skin_row[c] for c in COL_MISC_PLIES]

            deeper = find_neighbors(neighbor_org, prop_nodes, exclude={pbarl})
            deeper_pid = deeper[0] if deeper else None

            if deeper_pid is not None and prop_types.get(deeper_pid) != "PBARL":
                lookup_pid = deeper_pid
            else:
                lookup_pid = neighbor_org

            reverse = check_angle(bar_element, lookup_pid, bdf, tol=ANGLE_TOL_DEG)

            load_row = df_load[df_load[COL_LOAD_PID] == lookup_pid]
            a = load_row[COL_LOAD_A].max()
            b = load_row[COL_LOAD_B].max()
            if reverse:
                a, b = b, a

            rows.append((
                pbarl, T, W, mat, lookup_pid,
                *neighbor_plies,
                web_id, web_t,
            ))

        progress_cb(i, total)

    return pd.DataFrame(rows, columns=[
        "PBARL", "T", "W", "MAT", "PCOMP",
        "N1", "N2", "N3", "N4",
        "PSHELL", "T_PSHELL",
    ])


# ---------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------
class SimpleGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PBARL Neighbor Extractor")
        self.root.geometry("620x260")
        self.root.resizable(False, False)

        self.is_running = False

        pad = {"padx": 8, "pady": 6}

        self.bdf_path = tk.StringVar()
        self.load_path = tk.StringVar()
        self.misc_path = tk.StringVar()

        # --- Left image panel ---
        image_frame = tk.Frame(root, width=110)
        image_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(10, 0), pady=10)
        image_frame.pack_propagate(False)

        self.logo_image = None
        if IMAGE_PATH:
            try:
                self.logo_image = _load_logo_image(IMAGE_PATH, IMAGE_PANEL_SIZE)
                tk.Label(image_frame, image=self.logo_image).pack(expand=True)
            except Exception as e:
                print(f"[image] could not load '{IMAGE_PATH}': {e}")
                if not HAS_PIL:
                    print("[image] tip: install Pillow ('pip install pillow') "
                          "for JPEG support and better resizing")

        # --- Main content ---
        frame = tk.Frame(root)
        frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        frame.columnconfigure(1, weight=1)

        tk.Label(frame, text="BDF File:").grid(row=0, column=0, sticky="w", **pad)
        tk.Entry(frame, textvariable=self.bdf_path).grid(row=0, column=1, sticky="ew", **pad)
        tk.Button(frame, text="Browse", command=self.browse_bdf).grid(row=0, column=2, **pad)

        tk.Label(frame, text="Load CSV:").grid(row=1, column=0, sticky="w", **pad)
        tk.Entry(frame, textvariable=self.load_path).grid(row=1, column=1, sticky="ew", **pad)
        tk.Button(frame, text="Browse", command=self.browse_load).grid(row=1, column=2, **pad)

        tk.Label(frame, text="Misc Excel:").grid(row=2, column=0, sticky="w", **pad)
        tk.Entry(frame, textvariable=self.misc_path).grid(row=2, column=1, sticky="ew", **pad)
        tk.Button(frame, text="Browse", command=self.browse_misc).grid(row=2, column=2, **pad)

        self.progress = ttk.Progressbar(frame, orient="horizontal", mode="determinate")
        self.progress.grid(row=3, column=0, columnspan=3, sticky="ew", padx=8, pady=(16, 6))

        self.status_label = tk.Label(frame, text="Ready")
        self.status_label.grid(row=4, column=0, columnspan=3, sticky="w", padx=8)

        btn_frame = tk.Frame(frame)
        btn_frame.grid(row=5, column=0, columnspan=3, pady=(16, 0))
        self.run_btn = tk.Button(btn_frame, text="Run", width=12, command=self.start_run)
        self.run_btn.pack(side=tk.LEFT, padx=6)
        self.cancel_btn = tk.Button(btn_frame, text="Cancel", width=12, command=self.cancel_run,
                                     state=tk.DISABLED)
        self.cancel_btn.pack(side=tk.LEFT, padx=6)

    def browse_bdf(self):
        fn = filedialog.askopenfilename(filetypes=(("BDF files", "*.bdf"), ("All files", "*.*")))
        if fn:
            self.bdf_path.set(fn)

    def browse_load(self):
        fn = filedialog.askopenfilename(filetypes=(("CSV files", "*.csv"), ("All files", "*.*")))
        if fn:
            self.load_path.set(fn)

    def browse_misc(self):
        fn = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*")))
        if fn:
            self.misc_path.set(fn)

    def start_run(self):
        if not all([self.bdf_path.get(), self.load_path.get(), self.misc_path.get()]):
            messagebox.showwarning("Missing input", "Please select all three input files first.")
            return
        for label, path in (("BDF", self.bdf_path.get()), ("Load CSV", self.load_path.get()),
                            ("Misc Excel", self.misc_path.get())):
            if not os.path.exists(path):
                messagebox.showwarning("File not found", f"{label} file not found:\n{path}")
                return

        self.is_running = True
        self.run_btn.configure(state=tk.DISABLED)
        self.cancel_btn.configure(state=tk.NORMAL)
        self.progress["value"] = 0
        self.status_label.configure(text="Running...")

        thread = threading.Thread(target=self._run_worker, daemon=True)
        thread.start()

    def cancel_run(self):
        self.is_running = False
        self.status_label_update("Cancelling...")

    def status_label_update(self, text):
        self.root.after(0, lambda: self.status_label.configure(text=text))

    def set_progress(self, done, total):
        self.root.after(0, lambda: self.progress.configure(value=(done / total * 100) if total else 0))

    def _run_worker(self):
        try:
            bdf = read_bdf(self.bdf_path.get())
            df_load = pd.read_csv(self.load_path.get())
            df_misc = pd.read_excel(self.misc_path.get())

            df_out = run_extraction(
                bdf, df_load, df_misc,
                should_continue=lambda: self.is_running,
                progress_cb=self.set_progress,
            )

            out_path = os.path.join(os.getcwd(), "pbarl_neighbor_results.xlsx")
            df_out.to_excel(out_path, index=False)

            self.status_label_update(f"Done - {len(df_out)} rows -> {out_path}")

        except ExtractionCancelled:
            self.status_label_update("Cancelled")
        except Exception as e:
            self.status_label_update("Error")
            self.root.after(0, lambda: messagebox.showerror("Extraction Error", str(e)))
        finally:
            self.root.after(0, lambda: self.run_btn.configure(state=tk.NORMAL))
            self.root.after(0, lambda: self.cancel_btn.configure(state=tk.DISABLED))
            self.is_running = False


def main():
    root = tk.Tk()
    SimpleGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
