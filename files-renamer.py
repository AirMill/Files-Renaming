import os
import sys
import time
import subprocess
import tkinter as tk
from dataclasses import dataclass
from datetime import datetime
from tkinter import filedialog, messagebox
from tkinter import ttk
from typing import List, Optional, Tuple

# --- Drag & Drop (optional) ---
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:
    DND_FILES = None
    TkinterDnD = None

# --- EXIF (optional) ---
try:
    import piexif
    from PIL import Image
except Exception:
    piexif = None
    Image = None


# =========================
# Helpers
# =========================

@dataclass
class PlanItem:
    old_path: str
    new_path: str
    note: str = ""
    selected: bool = True


def safe_split_dnd(data: str, tk_root: tk.Tk) -> List[str]:
    try:
        return list(tk_root.tk.splitlist(data))
    except Exception:
        return data.split()


def sanitize_component(text: Optional[str]) -> str:
    if text is None:
        return ""
    text = text.strip()
    # Basic cross-platform filename safety
    for ch in ['/', '\\', ':', '*', '?', '"', '<', '>', '|', '\0']:
        text = text.replace(ch, " ")
    return " ".join(text.split())


def get_file_date_dt(path: str, mode: str) -> datetime:
    st = os.stat(path)
    if mode == "created":
        ts = getattr(st, "st_birthtime", st.st_ctime)  # macOS birthtime; else ctime
    else:
        ts = st.st_mtime
    return datetime.fromtimestamp(ts)


def can_write_exif(path: str) -> bool:
    ext = os.path.splitext(path)[1].lower()
    return ext in [".jpg", ".jpeg", ".tif", ".tiff"]


def exif_dt_string(dt: datetime) -> bytes:
    return dt.strftime("%Y:%m:%d %H:%M:%S").encode("utf-8")


def jpeg_seems_to_have_exif(path: str) -> bool:
    ext = os.path.splitext(path)[1].lower()
    if ext not in [".jpg", ".jpeg"]:
        return False
    try:
        with open(path, "rb") as f:
            head = f.read(256 * 1024)
        return b"Exif\x00\x00" in head
    except Exception:
        return False


def write_exif_fields(
    path: str,
    preserve_existing: bool,
    set_date: bool,
    date_value: Optional[datetime],
    set_author: bool,
    author: str
) -> Tuple[bool, str]:
    """
    Safe EXIF write:
      - If preserve_existing=True: load existing EXIF, update only selected tags.
      - If EXIF seems present but cannot be parsed: skip to avoid wiping metadata.
    Writes:
      - DateTimeOriginal, DateTimeDigitized, DateTime
      - Artist
    """
    if piexif is None or Image is None:
        return False, "EXIF libraries missing (install pillow + piexif)."

    if not can_write_exif(path):
        return False, "Not supported (JPEG/TIFF only)."

    # Verify readable image
    try:
        with Image.open(path) as img:
            img.verify()
    except Exception:
        return False, "Image unreadable."

    try:
        exif_dict = {"0th": {}, "Exif": {}, "GPS": {}, "1st": {}, "thumbnail": None}

        loaded_ok = False
        load_error = None
        try:
            loaded = piexif.load(path)
            if isinstance(loaded, dict):
                exif_dict = loaded
                loaded_ok = True
        except Exception as e:
            load_error = e

        if preserve_existing and (not loaded_ok):
            if jpeg_seems_to_have_exif(path):
                return False, f"Preserve ON: cannot parse EXIF ({load_error}); skipped."

        if set_author:
            a = author.strip()
            if a:
                exif_dict["0th"][piexif.ImageIFD.Artist] = a.encode("utf-8")

        if set_date and date_value is not None:
            dtb = exif_dt_string(date_value)
            exif_dict["Exif"][piexif.ExifIFD.DateTimeOriginal] = dtb
            exif_dict["Exif"][piexif.ExifIFD.DateTimeDigitized] = dtb
            exif_dict["0th"][piexif.ImageIFD.DateTime] = dtb

        exif_bytes_out = piexif.dump(exif_dict)
        piexif.insert(exif_bytes_out, path)
        return True, "EXIF updated"
    except Exception as e:
        return False, f"EXIF write failed: {e}"


def open_in_file_manager(path: str) -> None:
    folder = path if os.path.isdir(path) else os.path.dirname(path)
    try:
        if sys.platform.startswith("darwin"):
            subprocess.run(["open", folder], check=False)
        elif os.name == "nt":
            os.startfile(folder)  # type: ignore[attr-defined]
        else:
            subprocess.run(["xdg-open", folder], check=False)
    except Exception:
        pass


def unique_path(path: str, occupied: set) -> str:
    base, ext = os.path.splitext(path)
    candidate = path
    n = 1
    while candidate in occupied or os.path.exists(candidate):
        candidate = f"{base} ({n}){ext}"
        n += 1
    occupied.add(candidate)
    return candidate


# =========================
# Tooltips (small polish)
# =========================

class ToolTip:
    def __init__(self, widget, text: str):
        self.widget = widget
        self.text = text
        self.tip = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)

    def _show(self, _=None):
        if self.tip or not self.text:
            return
        try:
            x = self.widget.winfo_rootx() + 20
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
            self.tip = tw = tk.Toplevel(self.widget)
            tw.wm_overrideredirect(True)
            tw.wm_geometry(f"+{x}+{y}")
            label = ttk.Label(tw, text=self.text, padding=(10, 6))
            label.pack()
        except Exception:
            self.tip = None

    def _hide(self, _=None):
        if self.tip:
            try:
                self.tip.destroy()
            except Exception:
                pass
            self.tip = None


# =========================
# Main App
# =========================

class PhotoRenamerPro:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PhotoRenamer Pro")
        self.root.minsize(1040, 680)

        # Theme: Aqua on macOS, fallback otherwise
        style = ttk.Style()
        try:
            style.theme_use("aqua")
        except Exception:
            pass

        # Data
        self.files: List[str] = []
        self.last_undo_map: List[Tuple[str, str]] = []  # [(new, old)]
        self.current_plan: List[PlanItem] = []

        # Menubar + shortcuts
        self._build_menu()
        self._bind_shortcuts()

        # Top bar + drop zone
        self._build_top_toolbar()
        self._build_drop_zone()

        # Main Paned layout
        self.main_pane = ttk.Panedwindow(self.root, orient="horizontal")
        self.main_pane.pack(fill="both", expand=True, padx=12, pady=(0, 10))

        self.left = ttk.Frame(self.main_pane)
        self.right = ttk.Frame(self.main_pane)

        self.main_pane.add(self.left, weight=1)
        self.main_pane.add(self.right, weight=3)

        # Left: options
        self._build_options_rename(self.left)
        self._build_options_scope_sort(self.left)
        self._build_options_exif(self.left)

        # Right: preview
        self._build_preview(self.right)

        # Bottom status + progress
        self._build_status_bar()

        # DnD hookup
        if DND_FILES and hasattr(self.root, "drop_target_register"):
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<Drop>>", self.on_drop)
            self.drop_text.set("Drag & drop files or folders here — or use the toolbar.")
        else:
            self.drop_text.set("Drag & drop not available. Install: python3 -m pip install tkinterdnd2")

        self.refresh_preview()

    # --------------------
    # UI: Menu / Shortcuts
    # --------------------
    def _build_menu(self):
        menubar = tk.Menu(self.root)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Add Files…", command=self.add_files, accelerator="⌘O / Ctrl+O")
        file_menu.add_command(label="Add Folder…", command=self.add_folder, accelerator="⇧⌘O / Ctrl+Shift+O")
        file_menu.add_separator()
        file_menu.add_command(label="Open Containing Folder", command=self.open_selected_folder, accelerator="⌘K / Ctrl+K")
        file_menu.add_separator()
        file_menu.add_command(label="Clear List", command=self.clear, accelerator="⌘L / Ctrl+L")
        file_menu.add_separator()
        file_menu.add_command(label="Quit", command=self.root.quit, accelerator="⌘Q / Ctrl+Q")
        menubar.add_cascade(label="File", menu=file_menu)

        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="Select All", command=lambda: self.set_all_selected(True), accelerator="⌘A / Ctrl+A")
        edit_menu.add_command(label="Select None", command=lambda: self.set_all_selected(False), accelerator="⇧⌘A / Ctrl+Shift+A")
        edit_menu.add_separator()
        edit_menu.add_command(label="Undo Last Rename", command=self.undo, accelerator="⌘Z / Ctrl+Z")
        menubar.add_cascade(label="Edit", menu=edit_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self.about)
        menubar.add_cascade(label="Help", menu=help_menu)

        self.root.config(menu=menubar)

    def _bind_shortcuts(self):
        is_mac = sys.platform.startswith("darwin")
        cmd = "Command" if is_mac else "Control"

        self.root.bind_all(f"<{cmd}-o>", lambda e: self.add_files())
        self.root.bind_all(f"<Shift-{cmd}-O>", lambda e: self.add_folder())
        self.root.bind_all(f"<{cmd}-l>", lambda e: self.clear())
        self.root.bind_all(f"<{cmd}-z>", lambda e: self.undo())
        self.root.bind_all(f"<{cmd}-a>", lambda e: self.set_all_selected(True))
        self.root.bind_all(f"<Shift-{cmd}-A>", lambda e: self.set_all_selected(False))
        self.root.bind_all(f"<{cmd}-k>", lambda e: self.open_selected_folder())
        self.root.bind_all(f"<{cmd}-q>", lambda e: self.root.quit())

        # Rename shortcut
        self.root.bind_all(f"<{cmd}-r>", lambda e: self.rename())

    # --------------------
    # UI: Top + Drop
    # --------------------
    def _build_top_toolbar(self):
        bar = ttk.Frame(self.root, padding=(12, 10))
        bar.pack(fill="x")

        ttk.Button(bar, text="Add Files…", command=self.add_files).pack(side="left")
        ttk.Button(bar, text="Add Folder…", command=self.add_folder).pack(side="left", padx=(8, 0))
        ttk.Button(bar, text="Clear", command=self.clear).pack(side="left", padx=(8, 0))

        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=10)

        ttk.Button(bar, text="Rename", command=self.rename).pack(side="left")
        ttk.Button(bar, text="Undo", command=self.undo).pack(side="left", padx=(8, 0))
        ttk.Button(bar, text="Open Folder", command=self.open_selected_folder).pack(side="left", padx=(8, 0))

        self.count_var = tk.StringVar(value="0 files")
        ttk.Label(bar, textvariable=self.count_var).pack(side="right")

    def _build_drop_zone(self):
        frm = ttk.Frame(self.root, padding=(12, 0))
        frm.pack(fill="x")

        self.drop_text = tk.StringVar()
        lbl = ttk.Label(frm, textvariable=self.drop_text, anchor="center", justify="center", relief="ridge")
        lbl.pack(fill="x", pady=(6, 10), ipady=14)

    # --------------------
    # UI: Left options
    # --------------------
    def _build_options_rename(self, parent):
        box = ttk.Labelframe(parent, text="Rename", padding=(12, 10))
        box.pack(fill="x", pady=(0, 10))

        # Separator
        r0 = ttk.Frame(box)
        r0.pack(fill="x")
        ttk.Label(r0, text="Separator").pack(side="left")
        self.sep_var = tk.StringVar(value=" - ")
        ttk.Entry(r0, textvariable=self.sep_var, width=10).pack(side="left", padx=(8, 10))
        ttk.Label(r0, text="between components").pack(side="left")
        ToolTip(r0, "Used to join Name / Description / Date / Number.")

        # Name
        r1 = ttk.Frame(box)
        r1.pack(fill="x", pady=(10, 0))
        ttk.Label(r1, text="Name").pack(side="left")
        self.name_var = tk.StringVar(value="")
        ttk.Entry(r1, textvariable=self.name_var).pack(side="left", fill="x", expand=True, padx=(8, 0))
        ToolTip(r1, "Required. Becomes the first component of the new filename.")

        # Description
        r2 = ttk.Frame(box)
        r2.pack(fill="x", pady=(10, 0))
        self.use_original_desc_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            r2,
            text="Use original filename as description",
            variable=self.use_original_desc_var,
            command=self._toggle_desc
        ).pack(side="top", anchor="w")

        r2b = ttk.Frame(box)
        r2b.pack(fill="x", pady=(6, 0))
        ttk.Label(r2b, text="Description").pack(side="left")
        self.desc_var = tk.StringVar(value="")
        self.desc_entry = ttk.Entry(r2b, textvariable=self.desc_var, state="disabled")
        self.desc_entry.pack(side="left", fill="x", expand=True, padx=(8, 0))

        # Date + Number
        r3 = ttk.Frame(box)
        r3.pack(fill="x", pady=(10, 0))

        self.add_date_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(r3, text="Add date", variable=self.add_date_var).pack(side="left")

        self.date_mode_var = tk.StringVar(value="modified")
        ttk.Combobox(
            r3, textvariable=self.date_mode_var, state="readonly",
            values=["modified", "created"], width=10
        ).pack(side="left", padx=(8, 8))

        self.date_fmt_var = tk.StringVar(value="%Y-%m-%d")
        ttk.Entry(r3, textvariable=self.date_fmt_var, width=12).pack(side="left")
        ToolTip(r3, "Date uses file Modified/Created and the strftime format.")

        r4 = ttk.Frame(box)
        r4.pack(fill="x", pady=(8, 0))

        self.add_num_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(r4, text="Add number", variable=self.add_num_var).pack(side="left")

        ttk.Label(r4, text="Start").pack(side="left", padx=(10, 0))
        self.num_start_var = tk.StringVar(value="1")
        ttk.Entry(r4, textvariable=self.num_start_var, width=6).pack(side="left", padx=(6, 10))

        ttk.Label(r4, text="Zero pad").pack(side="left")
        self.num_pad_var = tk.StringVar(value="2")
        ttk.Entry(r4, textvariable=self.num_pad_var, width=6).pack(side="left", padx=(6, 0))

        for v in [self.sep_var, self.name_var, self.desc_var, self.add_date_var,
                  self.date_mode_var, self.date_fmt_var, self.add_num_var,
                  self.num_start_var, self.num_pad_var, self.use_original_desc_var]:
            try:
                v.trace_add("write", lambda *_: self.refresh_preview())
            except Exception:
                pass

    def _build_options_scope_sort(self, parent):
        box = ttk.Labelframe(parent, text="Scope & Sorting", padding=(12, 10))
        box.pack(fill="x", pady=(0, 10))

        self.recursive_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(box, text="Include subfolders (recursive)", variable=self.recursive_var).pack(anchor="w")

        r = ttk.Frame(box)
        r.pack(fill="x", pady=(8, 0))
        ttk.Label(r, text="Sort by").pack(side="left")
        self.sort_var = tk.StringVar(value="name")
        ttk.Combobox(
            r, textvariable=self.sort_var, state="readonly",
            values=["name", "date_modified", "date_created", "size"],
            width=14
        ).pack(side="left", padx=(8, 0))
        ttk.Button(r, text="Apply", command=self.apply_sort).pack(side="left", padx=(10, 0))
        ToolTip(r, "Sorting affects numbering order.")

        try:
            self.sort_var.trace_add("write", lambda *_: self.apply_sort())
        except Exception:
            pass

    def _build_options_exif(self, parent):
        box = ttk.Labelframe(parent, text="EXIF Tools", padding=(12, 10))
        box.pack(fill="x")

        lib_status = "OK" if (piexif and Image) else "Missing (install pillow + piexif)"
        self.exif_enable_var = tk.BooleanVar(value=False)

        r0 = ttk.Frame(box)
        r0.pack(fill="x")
        ttk.Checkbutton(
            r0, text="Enable EXIF edits (JPEG/TIFF)", variable=self.exif_enable_var,
            command=self._toggle_exif
        ).pack(side="left")
        ttk.Label(r0, text=f"Libs: {lib_status}").pack(side="right")

        self.exif_preserve_var = tk.BooleanVar(value=True)
        self.cb_exif_preserve = ttk.Checkbutton(
            box, text="Preserve ALL existing EXIF (keep other tags)",
            variable=self.exif_preserve_var
        )
        self.cb_exif_preserve.pack(anchor="w", pady=(6, 0))
        ToolTip(self.cb_exif_preserve, "If EXIF is present but unreadable, files will be skipped to avoid wiping metadata.")

        r1 = ttk.Frame(box)
        r1.pack(fill="x", pady=(8, 0))
        self.exif_set_date_var = tk.BooleanVar(value=True)
        self.cb_exif_date = ttk.Checkbutton(r1, text="Set EXIF date", variable=self.exif_set_date_var)
        self.cb_exif_date.pack(side="left")

        self.exif_date_mode_var = tk.StringVar(value="file_modified")
        self.exif_date_mode_combo = ttk.Combobox(
            r1, textvariable=self.exif_date_mode_var, state="readonly",
            values=["file_modified", "file_created", "custom"], width=14
        )
        self.exif_date_mode_combo.pack(side="left", padx=(8, 0))

        r2 = ttk.Frame(box)
        r2.pack(fill="x", pady=(6, 0))
        ttk.Label(r2, text="Custom date").pack(side="left")
        self.exif_custom_date_var = tk.StringVar(value="")
        self.exif_custom_entry = ttk.Entry(r2, textvariable=self.exif_custom_date_var)
        self.exif_custom_entry.pack(side="left", fill="x", expand=True, padx=(8, 0))
        ToolTip(r2, "Format: YYYY-MM-DD or YYYY-MM-DD HH:MM:SS")

        r3 = ttk.Frame(box)
        r3.pack(fill="x", pady=(8, 0))
        self.exif_set_author_var = tk.BooleanVar(value=False)
        self.cb_exif_author = ttk.Checkbutton(r3, text="Set EXIF author", variable=self.exif_set_author_var)
        self.cb_exif_author.pack(side="left")

        r4 = ttk.Frame(box)
        r4.pack(fill="x", pady=(6, 0))
        ttk.Label(r4, text="Author").pack(side="left")
        self.exif_author_var = tk.StringVar(value="")
        self.exif_author_entry = ttk.Entry(r4, textvariable=self.exif_author_var)
        self.exif_author_entry.pack(side="left", fill="x", expand=True, padx=(8, 0))
        ToolTip(r4, "Writes EXIF Artist tag (common for author/copyright workflows).")

        self._toggle_exif()

        for v in [self.exif_enable_var, self.exif_preserve_var, self.exif_set_date_var,
                  self.exif_date_mode_var, self.exif_custom_date_var,
                  self.exif_set_author_var, self.exif_author_var]:
            try:
                v.trace_add("write", lambda *_: self.refresh_preview())
            except Exception:
                pass

    # --------------------
    # UI: Preview table
    # --------------------
    def _build_preview(self, parent):
        box = ttk.Labelframe(parent, text="Preview", padding=(12, 10))
        box.pack(fill="both", expand=True)

        toolbar = ttk.Frame(box)
        toolbar.pack(fill="x", pady=(0, 8))

        ttk.Button(toolbar, text="Select All", command=lambda: self.set_all_selected(True)).pack(side="left")
        ttk.Button(toolbar, text="Select None", command=lambda: self.set_all_selected(False)).pack(side="left", padx=(8, 0))
        ttk.Button(toolbar, text="Toggle Selected", command=self.toggle_selected_rows).pack(side="left", padx=(8, 0))

        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=10)

        ttk.Button(toolbar, text="Refresh", command=self.refresh_preview).pack(side="left")

        self.tree = ttk.Treeview(
            box,
            columns=("sel", "old", "new", "folder", "notes"),
            show="headings",
            height=16
        )
        self.tree.heading("sel", text="✓")
        self.tree.heading("old", text="Current name")
        self.tree.heading("new", text="New name")
        self.tree.heading("folder", text="Folder")
        self.tree.heading("notes", text="Notes")

        self.tree.column("sel", width=40, anchor="center", stretch=False)
        self.tree.column("old", width=260, anchor="w")
        self.tree.column("new", width=300, anchor="w")
        self.tree.column("folder", width=200, anchor="w")
        self.tree.column("notes", width=240, anchor="w")

        self.tree.pack(side="left", fill="both", expand=True)

        sb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        sb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=sb.set)

        # Nice-to-have row striping (light theme)
        try:
            self.tree.tag_configure("odd", background="#f7f7f7")
            self.tree.tag_configure("even", background="white")
        except Exception:
            pass

        # Click on "sel" to toggle
        self.tree.bind("<Button-1>", self._on_tree_click)

    def _on_tree_click(self, event):
        # If user clicks on selection column, toggle that row
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        if col != "#1":  # "sel" is first displayed column
            return
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return
        idx = self._row_id_to_index(row_id)
        if idx is None:
            return
        self.current_plan[idx].selected = not self.current_plan[idx].selected
        self._update_row_sel(row_id, self.current_plan[idx].selected)
        self._update_selected_count()

    def _row_id_to_index(self, row_id: str) -> Optional[int]:
        # We store index in iid
        try:
            return int(row_id)
        except Exception:
            return None

    def _update_row_sel(self, row_id: str, selected: bool):
        vals = list(self.tree.item(row_id, "values"))
        vals[0] = "✓" if selected else ""
        self.tree.item(row_id, values=vals)

    # --------------------
    # UI: Status bar
    # --------------------
    def _build_status_bar(self):
        bar = ttk.Frame(self.root, padding=(12, 10))
        bar.pack(fill="x")

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(bar, textvariable=self.status_var).pack(side="left")

        self.selected_var = tk.StringVar(value="")
        ttk.Label(bar, textvariable=self.selected_var).pack(side="left", padx=(12, 0))

        self.progress = ttk.Progressbar(bar, mode="determinate", length=220)
        self.progress.pack(side="right")
        self.progress.pack_forget()

    # --------------------
    # Actions: About
    # --------------------
    def about(self):
        exif_ok = "Yes" if (piexif and Image) else "No (install pillow + piexif)"
        dnd_ok = "Yes" if DND_FILES and hasattr(self.root, "drop_target_register") else "No (install tkinterdnd2)"
        messagebox.showinfo(
            "About PhotoRenamer Pro",
            "PhotoRenamer Pro\n\n"
            "• Batch renaming with preview + undo\n"
            "• Recursive folder import + sorting\n"
            "• Optional EXIF edits (date, author) with preservation\n\n"
            f"Drag & drop: {dnd_ok}\n"
            f"EXIF libs: {exif_ok}\n"
        )

    # --------------------
    # Intake: Add / Drop
    # --------------------
    def add_files(self):
        paths = filedialog.askopenfilenames(title="Select files")
        if paths:
            self._add_paths(list(paths))

    def add_folder(self):
        folder = filedialog.askdirectory(title="Select folder")
        if folder:
            self._add_paths([folder])

    def on_drop(self, event):
        paths = safe_split_dnd(event.data, self.root)
        self._add_paths(paths)

    def _add_paths(self, paths: List[str]):
        recursive = bool(self.recursive_var.get())
        added = 0

        for p in paths:
            p = p.strip()
            if not p:
                continue

            if os.path.isdir(p):
                if recursive:
                    for root, _, files in os.walk(p):
                        for name in files:
                            fp = os.path.join(root, name)
                            if os.path.isfile(fp) and fp not in self.files:
                                self.files.append(fp)
                                added += 1
                else:
                    try:
                        for name in os.listdir(p):
                            fp = os.path.join(p, name)
                            if os.path.isfile(fp) and fp not in self.files:
                                self.files.append(fp)
                                added += 1
                    except Exception:
                        pass
            elif os.path.isfile(p):
                if p not in self.files:
                    self.files.append(p)
                    added += 1

        self.apply_sort(refresh=False)
        self.status_var.set(f"Added {added} item(s).")
        self.refresh_preview()

    def clear(self):
        self.files = []
        self.last_undo_map = []
        self.current_plan = []
        self.status_var.set("Cleared.")
        self.refresh_preview()

    # --------------------
    # Sorting
    # --------------------
    def apply_sort(self, refresh: bool = True):
        key = self.sort_var.get()

        def safe_stat(path: str):
            try:
                return os.stat(path)
            except Exception:
                return None

        if key == "name":
            self.files.sort(key=lambda x: os.path.basename(x).lower())
        elif key == "size":
            self.files.sort(key=lambda x: ((safe_stat(x).st_size if safe_stat(x) else -1), os.path.basename(x).lower()))
        elif key == "date_created":
            self.files.sort(key=lambda x: get_file_date_dt(x, "created").timestamp())
        else:  # date_modified
            self.files.sort(key=lambda x: get_file_date_dt(x, "modified").timestamp())

        if refresh:
            self.refresh_preview()

    # --------------------
    # Rename plan building
    # --------------------
    def _toggle_desc(self):
        self.desc_entry.configure(state="disabled" if self.use_original_desc_var.get() else "normal")
        self.refresh_preview()

    def _toggle_exif(self):
        enabled = bool(self.exif_enable_var.get()) and (piexif is not None and Image is not None)
        state = "normal" if enabled else "disabled"
        for w in [self.cb_exif_preserve, self.cb_exif_date, self.exif_date_mode_combo,
                  self.exif_custom_entry, self.cb_exif_author, self.exif_author_entry]:
            w.configure(state=state)
        self.refresh_preview()

    def _parse_int(self, s: str, default: int) -> int:
        try:
            return int(str(s).strip())
        except Exception:
            return default

    def _compute_exif_date_for_file(self, path: str) -> Optional[datetime]:
        mode = self.exif_date_mode_var.get()
        if mode == "file_created":
            return get_file_date_dt(path, "created")
        if mode == "file_modified":
            return get_file_date_dt(path, "modified")

        raw = (self.exif_custom_date_var.get() or "").strip()
        if not raw:
            return None
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
            try:
                dt = datetime.strptime(raw, fmt)
                if fmt == "%Y-%m-%d":
                    dt = dt.replace(hour=0, minute=0, second=0)
                return dt
            except Exception:
                pass
        return None

    def build_plan(self) -> List[PlanItem]:
        sep = self.sep_var.get() if self.sep_var.get() is not None else " "
        name = sanitize_component(self.name_var.get())

        use_original_desc = bool(self.use_original_desc_var.get())
        typed_desc = sanitize_component(self.desc_var.get())

        add_date = bool(self.add_date_var.get())
        date_mode = self.date_mode_var.get()
        date_fmt = self.date_fmt_var.get() or "%Y-%m-%d"

        add_num = bool(self.add_num_var.get())
        start = self._parse_int(self.num_start_var.get(), 1)
        pad = max(0, self._parse_int(self.num_pad_var.get(), 2))

        exif_enabled = bool(self.exif_enable_var.get()) and (piexif is not None and Image is not None)
        exif_set_date = exif_enabled and bool(self.exif_set_date_var.get())
        exif_set_author = exif_enabled and bool(self.exif_set_author_var.get())
        preserve = bool(self.exif_preserve_var.get())
        author = (self.exif_author_var.get() or "").strip()

        occupied = set()
        plan: List[PlanItem] = []
        n = start

        for path in self.files:
            dirname = os.path.dirname(path)
            base = os.path.basename(path)
            stem, ext = os.path.splitext(base)

            parts: List[str] = []
            notes: List[str] = []

            if name:
                parts.append(name)
            else:
                notes.append("Missing Name")

            if use_original_desc:
                d = sanitize_component(stem)
                if d:
                    parts.append(d)
            else:
                if typed_desc:
                    parts.append(typed_desc)

            if add_date:
                try:
                    dt = get_file_date_dt(path, date_mode)
                    parts.append(dt.strftime(date_fmt))
                except Exception:
                    parts.append("DATE_ERR")
                    notes.append("Date error")

            if add_num:
                num_str = str(n).zfill(pad) if pad > 0 else str(n)
                parts.append(num_str)
                n += 1

            new_stem = sep.join([p for p in parts if p != ""])
            if not new_stem:
                new_stem = stem

            new_path = unique_path(os.path.join(dirname, new_stem + ext), occupied)

            if exif_enabled and (exif_set_date or exif_set_author):
                if can_write_exif(path):
                    if preserve:
                        notes.append("EXIF preserve ON")
                    if exif_set_date:
                        dtv = self._compute_exif_date_for_file(path)
                        notes.append("EXIF date set" if dtv else "EXIF date invalid")
                    if exif_set_author:
                        notes.append("EXIF author set" if author else "EXIF author empty")
                else:
                    notes.append("EXIF skipped (format)")

            plan.append(PlanItem(path, new_path, "; ".join(notes), True))

        return plan

    # --------------------
    # Preview refresh
    # --------------------
    def refresh_preview(self):
        if not hasattr(self, "tree"):
            return

        self.count_var.set(f"{len(self.files)} file(s)")

        # Build plan
        self.current_plan = self.build_plan()

        # Clear view
        for item in self.tree.get_children():
            self.tree.delete(item)

        if not self.files:
            self.status_var.set("Ready (no files).")
            self.selected_var.set("")
            return

        # Populate
        for idx, p in enumerate(self.current_plan):
            old = os.path.basename(p.old_path)
            new = os.path.basename(p.new_path)
            folder = os.path.basename(os.path.dirname(p.old_path))
            sel = "✓" if p.selected else ""
            tag = "even" if idx % 2 == 0 else "odd"
            # store idx as iid for easy mapping
            self.tree.insert("", "end", iid=str(idx), values=(sel, old, new, folder, p.note), tags=(tag,))

        self._update_selected_count()
        self.status_var.set("Preview updated.")

    def _update_selected_count(self):
        total = len(self.current_plan)
        selected = sum(1 for p in self.current_plan if p.selected)
        self.selected_var.set(f"Selected: {selected}/{total}")

    def set_all_selected(self, selected: bool):
        if not self.current_plan:
            return
        for i, p in enumerate(self.current_plan):
            p.selected = selected
            row_id = str(i)
            if self.tree.exists(row_id):
                self._update_row_sel(row_id, selected)
        self._update_selected_count()

    def toggle_selected_rows(self):
        items = self.tree.selection()
        if not items:
            return
        for row_id in items:
            idx = self._row_id_to_index(row_id)
            if idx is None or idx < 0 or idx >= len(self.current_plan):
                continue
            self.current_plan[idx].selected = not self.current_plan[idx].selected
            self._update_row_sel(row_id, self.current_plan[idx].selected)
        self._update_selected_count()

    # --------------------
    # Open folder
    # --------------------
    def open_selected_folder(self):
        items = self.tree.selection()
        if items:
            idx = self._row_id_to_index(items[0])
            if idx is not None and 0 <= idx < len(self.current_plan):
                open_in_file_manager(self.current_plan[idx].old_path)
                return
        # fallback: open folder of first file
        if self.files:
            open_in_file_manager(self.files[0])

    # --------------------
    # Rename + Undo
    # --------------------
    def rename(self):
        if not self.files:
            messagebox.showwarning("No files", "Add files or a folder first.")
            return

        if not sanitize_component(self.name_var.get()):
            messagebox.showwarning("Name required", "Please enter the Name (first component).")
            return

        plan = self.current_plan if self.current_plan else self.build_plan()
        effective = [p for p in plan if p.selected and p.old_path != p.new_path]

        if not effective:
            messagebox.showinfo("Nothing to do", "No selected files would be renamed.")
            return

        exif_enabled = bool(self.exif_enable_var.get()) and (piexif is not None and Image is not None)
        exif_set_date = exif_enabled and bool(self.exif_set_date_var.get())
        exif_set_author = exif_enabled and bool(self.exif_set_author_var.get())
        preserve = bool(self.exif_preserve_var.get())
        author = (self.exif_author_var.get() or "").strip()

        exif_warning = ""
        if exif_enabled and (exif_set_date or exif_set_author):
            exif_warning = "\n\nNote: EXIF edits are not undoable by Undo Rename."

        if not messagebox.askyesno("Confirm Rename", f"Rename {len(effective)} file(s)?{exif_warning}"):
            return

        # Progress
        self.progress.pack(side="right")
        self.progress["maximum"] = len(effective)
        self.progress["value"] = 0
        self.root.update_idletasks()

        timestamp = str(int(time.time()))
        temp_map: List[Tuple[str, str, str]] = []  # (old, temp, final)
        self.last_undo_map = []

        # Phase 1 + 2 rename
        try:
            for i, p in enumerate(effective):
                old = p.old_path
                final = p.new_path
                d = os.path.dirname(old)
                _, ext = os.path.splitext(final)
                temp = os.path.join(d, f".__photorenamer_tmp_{timestamp}_{i}{ext}")
                os.rename(old, temp)
                temp_map.append((old, temp, final))
                self.progress["value"] = i + 0.25
                if i % 25 == 0:
                    self.root.update_idletasks()

            for i, (old, temp, final) in enumerate(temp_map):
                os.rename(temp, final)
                self.last_undo_map.append((final, old))
                self.progress["value"] = i + 1
                if i % 25 == 0:
                    self.root.update_idletasks()
        except Exception as e:
            # rollback best-effort
            for old, temp, final in reversed(temp_map):
                try:
                    if os.path.exists(final) and not os.path.exists(old):
                        os.rename(final, old)
                    elif os.path.exists(temp) and not os.path.exists(old):
                        os.rename(temp, old)
                except Exception:
                    pass
            self.progress.pack_forget()
            messagebox.showerror("Rename failed", str(e))
            self.refresh_preview()
            return

        # EXIF pass (optional)
        exif_results = {"updated": 0, "skipped": 0, "failed": 0}
        exif_fail_msgs: List[str] = []
        if exif_enabled and (exif_set_date or exif_set_author):
            for final, old in self.last_undo_map:
                if not can_write_exif(final):
                    exif_results["skipped"] += 1
                    continue
                dt_val = self._compute_exif_date_for_file(old) if exif_set_date else None
                ok, msg = write_exif_fields(
                    final,
                    preserve_existing=preserve,
                    set_date=exif_set_date,
                    date_value=dt_val,
                    set_author=exif_set_author,
                    author=author
                )
                if ok:
                    exif_results["updated"] += 1
                else:
                    exif_results["failed"] += 1
                    if len(exif_fail_msgs) < 4:
                        exif_fail_msgs.append(f"{os.path.basename(final)}: {msg}")

        # Update internal file list with renamed paths (selected only)
        lookup = {p.old_path: p.new_path for p in effective}
        self.files = [lookup.get(f, f) for f in self.files]
        self.apply_sort(refresh=False)
        self.progress.pack_forget()
        self.refresh_preview()

        extra = ""
        if exif_enabled and (exif_set_date or exif_set_author):
            extra = f"\n\nEXIF: updated {exif_results['updated']}, skipped {exif_results['skipped']}, failed {exif_results['failed']}."
            if exif_fail_msgs:
                extra += "\nSome EXIF notes:\n- " + "\n- ".join(exif_fail_msgs)

        messagebox.showinfo("Success", f"Renamed {len(effective)} file(s). Undo is available for renames in this session.{extra}")
        self.status_var.set(f"Renamed {len(effective)} file(s).")

    def undo(self):
        if not self.last_undo_map:
            messagebox.showinfo("Undo", "Nothing to undo (only last rename in this session).")
            return

        if not messagebox.askyesno("Undo Rename", f"Undo last rename for {len(self.last_undo_map)} file(s)?"):
            return

        self.progress.pack(side="right")
        self.progress["maximum"] = len(self.last_undo_map)
        self.progress["value"] = 0
        self.root.update_idletasks()

        timestamp = str(int(time.time()))
        temp_map: List[Tuple[str, str, str]] = []  # (final, temp, old)

        try:
            for i, (final, old) in enumerate(self.last_undo_map):
                d = os.path.dirname(final)
                _, ext = os.path.splitext(final)
                temp = os.path.join(d, f".__photorenamer_undo_tmp_{timestamp}_{i}{ext}")
                if os.path.exists(final):
                    os.rename(final, temp)
                    temp_map.append((final, temp, old))
                self.progress["value"] = i + 0.5
                if i % 25 == 0:
                    self.root.update_idletasks()

            for i, (final, temp, old) in enumerate(temp_map):
                os.rename(temp, old)
                self.progress["value"] = i + 1
                if i % 25 == 0:
                    self.root.update_idletasks()
        except Exception as e:
            self.progress.pack_forget()
            messagebox.showerror("Undo failed", str(e))
            return

        lookup = {final: old for (final, old) in self.last_undo_map}
        self.files = [lookup.get(f, f) for f in self.files]
        self.apply_sort(refresh=False)
        self.last_undo_map = []
        self.progress.pack_forget()
        self.refresh_preview()
        self.status_var.set("Undo complete.")
        messagebox.showinfo("Undo", "Undo complete (renames only).")


# =========================
# Entrypoint
# =========================

def main():
    root = TkinterDnD.Tk() if TkinterDnD else tk.Tk()
    try:
        root.tk.call("tk", "scaling", 1.1)
    except Exception:
        pass
    PhotoRenamerPro(root)
    root.mainloop()


if __name__ == "__main__":
    main()