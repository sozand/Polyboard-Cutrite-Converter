import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
import pandas as pd
import re
from datetime import datetime
import json
import uuid
import sys
from Detect_Processes_In_mpr_file_ import map_and_count_mpr_processes

try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

# Dark theme palette
DARK_BG = "#1e1e1e"
DARK_FG = "#e0e0e0"
DARK_ACCENT = "#3a8fb7"
DARK_ENTRY = "#252526"
DARK_HILIGHT = "#2d2d30"


# Fixed text block to remove from MPR files
# Overall flow (high level):
# main() -> PolyboardProductionGUI creates ttk.Notebook with two tabs
#   Tab 1 (MPRProcessorTab): user picks folder -> preview finds .mpr files
#   containing the fixed block -> process removes block with backup/optional confirm
#   Tab 2 (CutlistGeneratorTab): user picks convention Excel + cutlist CSV ->
#   load -> process converts grain, maps component/edge codes, adds face name ->
#   preview -> export CSV
MPR_TEXT_TO_REMOVE = '''<139 \\Komponente\\
IN="ZP500_FR.mpr"
KAT="Composant"
MNM="Composant"'''

# Column layout for cutlist CSVs (files have no headers; we map them explicitly)
CUTLIST_COLUMNS = [
    "Reference",
    "Material",
    "Height_Net",
    "Width_Net",
    "Quantity",
    "Grain_Direction",
    "Right_Edge",
    "Left_Edge",
    "Bottom_Edge",
    "Top_Edge",
    "Cutting_List_Number",
    "Project",
    "Cabinet",
    "Height_Overall",
    "Width_Overall",
    "Tooling_File_First_Face",
    "Edging_Diagram",
    "Thickness",
    "Face",
]

CONVENTION_COLUMNS = [
    "Component",
    "Face_1",
    "Face_2",
    "Edge_0",
    "Edge_1",
    "Edge_2_no_connect",
    "Edge_2_connect",
    "Edge_3",
    "Edge_4",
]


def get_app_base_dir() -> Path:
    """Directory where the script/exe lives (portable-friendly)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


CONFIG_FILENAME = "polyboard_config.json"


def get_config_path() -> Path:
    return get_app_base_dir() / CONFIG_FILENAME


def apply_dark_theme(root: tk.Tk):
    """Apply a simple dark theme to ttk and Tk widgets."""
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    root.configure(bg=DARK_BG)
    style.configure(".", background=DARK_BG, foreground=DARK_FG)
    style.configure("TFrame", background=DARK_BG)
    style.configure("TLabel", background=DARK_BG, foreground=DARK_FG)
    style.configure("TButton", background=DARK_HILIGHT, foreground=DARK_FG, padding=5)
    style.map("TButton", background=[("active", DARK_ACCENT)], foreground=[("disabled", "#888888")])
    style.configure("TCheckbutton", background=DARK_BG, foreground=DARK_FG)
    style.configure("TNotebook", background=DARK_BG, foreground=DARK_FG)
    style.configure("TNotebook.Tab", background=DARK_HILIGHT, foreground=DARK_FG, padding=(10, 5))
    style.map("TNotebook.Tab", background=[("selected", DARK_ACCENT)], foreground=[("selected", DARK_FG)])
    style.configure("TEntry", fieldbackground=DARK_ENTRY, foreground=DARK_FG, insertcolor=DARK_FG)
    style.configure("Treeview", background=DARK_ENTRY, fieldbackground=DARK_ENTRY, foreground=DARK_FG)
    style.configure("Treeview.Heading", background=DARK_HILIGHT, foreground=DARK_FG)
    style.map("Treeview", background=[("selected", DARK_ACCENT)])


class ConventionEditorDialog:
    """Modal dialog to edit convention data (CRUD) with import/export."""

    def __init__(self, parent, initial_df: pd.DataFrame, json_path: Path, on_save_callback, edge_dir: Path = None):
        self.parent = parent
        self.json_path = json_path
        self.on_save_callback = on_save_callback
        self.data_df = initial_df.copy() if initial_df is not None else pd.DataFrame(columns=CONVENTION_COLUMNS)
        self.image_thumbs = []
        self.image_full = {}
        self.edge_dir = edge_dir if edge_dir is not None else Path(__file__).resolve().parent / "Edge_Diagram_Ref"

        self.window = tk.Toplevel(parent)
        self.window.title("Convention Editor")
        self.window.geometry("1200x700")
        self.window.transient(parent)
        self.window.grab_set()

        apply_dark_theme(self.window)

        self._build_ui()
        self._populate_tree()
        self._load_images_panel()

    def _build_ui(self):
        top_frame = ttk.Frame(self.window, padding="8")
        top_frame.pack(fill=tk.BOTH, expand=True)

        btn_frame = ttk.Frame(top_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 8))

        ttk.Button(btn_frame, text="Add Row", command=self._add_row).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_frame, text="Update Row", command=self._update_row).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_frame, text="Delete Row", command=self._delete_row).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_frame, text="Save All (JSON)", command=self._save_all).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_frame, text="Import Excel", command=self._import_excel).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_frame, text="Export Excel", command=self._export_excel).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_frame, text="Close", command=self.window.destroy).pack(side=tk.RIGHT, padx=4)

        tree_frame = ttk.Frame(top_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(tree_frame, columns=CONVENTION_COLUMNS, show="headings", height=15)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        for col in CONVENTION_COLUMNS:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=130, anchor="w")

        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        form = ttk.LabelFrame(top_frame, text="Edit Selected / New Row", padding="8")
        form.pack(fill=tk.X, pady=(8, 0))

        self.entries = {}
        for i, col in enumerate(CONVENTION_COLUMNS):
            row_frame = ttk.Frame(form)
            row_frame.pack(fill=tk.X, pady=2)
            ttk.Label(row_frame, text=col + ":").pack(side=tk.LEFT, padx=(0, 6))
            ent = ttk.Entry(row_frame, width=60)
            ent.pack(side=tk.LEFT, fill=tk.X, expand=True)
            self.entries[col] = ent

        # Image panel
        img_frame = ttk.LabelFrame(top_frame, text="Edge Diagram Reference (Edge_Diagram_Ref)", padding="8")
        img_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        self.img_canvas = tk.Canvas(img_frame, bg=DARK_BG, highlightthickness=0)
        self.img_vsb = ttk.Scrollbar(img_frame, orient="vertical", command=self.img_canvas.yview)
        self.img_hsb = ttk.Scrollbar(img_frame, orient="horizontal", command=self.img_canvas.xview)
        self.img_canvas.configure(yscrollcommand=self.img_vsb.set, xscrollcommand=self.img_hsb.set)

        self.img_inner = ttk.Frame(self.img_canvas)
        self.img_canvas.create_window((0, 0), window=self.img_inner, anchor="nw")

        self.img_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.img_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.img_hsb.pack(side=tk.BOTTOM, fill=tk.X)

        self.img_inner.bind("<Configure>", lambda e: self.img_canvas.configure(scrollregion=self.img_canvas.bbox("all")))

    def _populate_tree(self):
        self.tree.delete(*self.tree.get_children())
        df = self.data_df
        for _, row in df.iterrows():
            values = [row.get(col, "") for col in CONVENTION_COLUMNS]
            self.tree.insert("", "end", values=values)

    def _load_images_panel(self):
        # Clear previous thumbs
        self.image_thumbs.clear()
        for child in list(self.img_inner.winfo_children()):
            child.destroy()

        ref_dir = self.edge_dir
        if not ref_dir.exists():
            ttk.Label(self.img_inner, text=f"Edge_Diagram_Ref folder not found: {ref_dir}", foreground=DARK_FG, background=DARK_BG).pack(anchor="w", padx=4, pady=2)
            return
        files = sorted([p for p in ref_dir.iterdir() if p.is_file() and p.suffix.lower() in [".png", ".jpg", ".jpeg", ".gif", ".bmp"]], key=lambda p: p.name)
        if not files:
            ttk.Label(self.img_inner, text="No images in Edge_Diagram_Ref.", foreground=DARK_FG, background=DARK_BG).pack(anchor="w", padx=4, pady=2)
            return
        if not PIL_AVAILABLE:
            ttk.Label(self.img_inner, text="Pillow not installed; cannot load images.", foreground=DARK_FG, background=DARK_BG).pack(anchor="w", padx=4, pady=2)
            for p in files:
                ttk.Label(self.img_inner, text=p.name, foreground=DARK_FG, background=DARK_BG).pack(anchor="w", padx=4, pady=2)
            return

        for p in files:
            try:
                img = Image.open(p)
                img.thumbnail((300, 300))
                tkimg = ImageTk.PhotoImage(img)
                self.image_thumbs.append(tkimg)
                self.image_full[p.name] = Image.open(p)  # store full image for zoom
                frame = ttk.Frame(self.img_inner)
                frame.pack(anchor="w", padx=4, pady=4, fill=tk.X)
                ttk.Label(frame, text=p.name).pack(anchor="w")
                lbl = tk.Label(frame, image=tkimg, bg=DARK_BG, cursor="hand2")
                lbl.pack(anchor="w")
                lbl.bind("<Button-1>", lambda e, name=p.name: self._open_image_popup(name))
            except Exception as e:
                ttk.Label(self.img_inner, text=f"{p.name} (error: {e})", foreground=DARK_FG, background=DARK_BG).pack(anchor="w", padx=4, pady=2)

    def _open_image_popup(self, name: str):
        if not PIL_AVAILABLE or name not in self.image_full:
            return
        img = self.image_full[name]
        popup = tk.Toplevel(self.window)
        popup.title(name)
        popup.geometry("900x600")
        popup.transient(self.window)
        apply_dark_theme(popup)

        canvas = tk.Canvas(popup, bg=DARK_BG, highlightthickness=0)
        vsb = ttk.Scrollbar(popup, orient="vertical", command=canvas.yview)
        hsb = ttk.Scrollbar(popup, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        popup.rowconfigure(0, weight=1)
        popup.columnconfigure(0, weight=1)

        # Resize to fit popup width while keeping aspect ratio
        def render_image():
            width = canvas.winfo_width()
            height = canvas.winfo_height()
            if width < 10 or height < 10:
                popup.after(50, render_image)
                return
            img_copy = img.copy()
            img_copy.thumbnail((max(width-20, 100), max(height-20, 100)))
            tkimg = ImageTk.PhotoImage(img_copy)
            canvas.image = tkimg  # keep reference
            canvas.delete("all")
            canvas.create_image(0, 0, anchor="nw", image=tkimg)
            canvas.config(scrollregion=canvas.bbox("all"))

        canvas.bind("<Configure>", lambda e: render_image())
        render_image()

    def _get_form_data(self) -> dict:
        return {col: self.entries[col].get().strip() for col in CONVENTION_COLUMNS}

    def _on_select(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        values = self.tree.item(sel[0], "values")
        for col, val in zip(CONVENTION_COLUMNS, values):
            self.entries[col].delete(0, tk.END)
            self.entries[col].insert(0, val)

    def _validate_unique_component(self, component: str, exclude_iid=None) -> bool:
        component_upper = component.upper()
        for iid in self.tree.get_children():
            if iid == exclude_iid:
                continue
            vals = self.tree.item(iid, "values")
            if vals and str(vals[0]).strip().upper() == component_upper:
                return False
        return True

    def _add_row(self):
        data = self._get_form_data()
        component = data.get("Component", "")
        if not component:
            messagebox.showwarning("Validation", "Component is required.")
            return
        if not self._validate_unique_component(component):
            messagebox.showwarning("Validation", "Component must be unique.")
            return
        values = [data.get(col, "") for col in CONVENTION_COLUMNS]
        self.tree.insert("", "end", values=values)

    def _update_row(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("No Selection", "Select a row to update.")
            return
        iid = sel[0]
        data = self._get_form_data()
        component = data.get("Component", "")
        if not component:
            messagebox.showwarning("Validation", "Component is required.")
            return
        if not self._validate_unique_component(component, exclude_iid=iid):
            messagebox.showwarning("Validation", "Component must be unique.")
            return
        values = [data.get(col, "") for col in CONVENTION_COLUMNS]
        self.tree.item(iid, values=values)

    def _delete_row(self):
        sel = self.tree.selection()
        if not sel:
            return
        self.tree.delete(sel[0])

    def _tree_to_df(self) -> pd.DataFrame:
        rows = []
        for iid in self.tree.get_children():
            vals = self.tree.item(iid, "values")
            rows.append({col: vals[i] if i < len(vals) else "" for i, col in enumerate(CONVENTION_COLUMNS)})
        return pd.DataFrame(rows, columns=CONVENTION_COLUMNS)

    def _save_json(self, df: pd.DataFrame):
        try:
            self.json_path.write_text(df.to_json(orient="records", force_ascii=False, indent=2), encoding="utf-8")
            self._log_status(f"Saved JSON: {self.json_path}")
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save JSON:\n{e}")

    def _log_status(self, msg: str):
        # minimal logging to console; dialog has no status box
        print(msg)

    def _save_all(self):
        df = self._tree_to_df()
        # validation: unique Component
        comps = df["Component"].fillna("").str.strip().str.upper()
        if (comps == "").any():
            messagebox.showwarning("Validation", "All rows must have Component.")
            return
        if comps.duplicated().any():
            messagebox.showwarning("Validation", "Component values must be unique.")
            return
        self.data_df = df
        self._save_json(df)
        if self.on_save_callback:
            self.on_save_callback(df)
        messagebox.showinfo("Saved", "Convention data saved.")

    def _import_excel(self):
        path = filedialog.askopenfilename(
            parent=self.window,
            title="Import Convention from Excel",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            df = pd.read_excel(path)
            missing = [c for c in CONVENTION_COLUMNS if c not in df.columns]
            if missing:
                messagebox.showerror("Import Error", f"Missing columns: {missing}")
                return
            df = df[CONVENTION_COLUMNS]
            self.data_df = df
            self._populate_tree()
            messagebox.showinfo("Import", "Imported convention from Excel.")
        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to import Excel:\n{e}")

    def _export_excel(self):
        path = filedialog.asksaveasfilename(
            parent=self.window,
            title="Export Convention to Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            df = self._tree_to_df()
            df.to_excel(path, index=False)
            messagebox.showinfo("Export", f"Exported to {path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export Excel:\n{e}")

class MPRProcessorTab:
    """Tab 1: MPR File Processor"""
    
    def __init__(self, parent):
        self.frame = ttk.Frame(parent, padding="10")
        self.frame.pack(fill=tk.BOTH, expand=True)
        
        self.project_folder = tk.StringVar()
        self.confirm_each = tk.BooleanVar(value=False)
        self.status_text = None
        
        self._create_widgets()
    
    def _create_widgets(self):
        # Title
        title = ttk.Label(self.frame, text="MPR File Processor", font=("Arial", 14, "bold"))
        title.pack(pady=(0, 10))
        
        # Folder selection
        folder_frame = ttk.Frame(self.frame)
        folder_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(folder_frame, text="Project Folder:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(folder_frame, textvariable=self.project_folder, width=50).pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
        ttk.Button(folder_frame, text="Browse", command=self._select_folder).pack(side=tk.LEFT)
        
        # Text to remove info
        info_frame = ttk.LabelFrame(self.frame, text="Text Block to Remove", padding="10")
        info_frame.pack(fill=tk.X, pady=10)
        
        info_text = tk.Text(info_frame, height=5, width=70, wrap=tk.WORD, state=tk.DISABLED)
        info_text.pack(fill=tk.X)
        info_text.config(state=tk.NORMAL)
        info_text.insert("1.0", MPR_TEXT_TO_REMOVE)
        info_text.config(state=tk.DISABLED)
        info_text.configure(bg=DARK_ENTRY, fg=DARK_FG, insertbackground=DARK_FG)
        
        # Options
        options_frame = ttk.Frame(self.frame)
        options_frame.pack(fill=tk.X, pady=5)
        
        ttk.Checkbutton(options_frame, text="Confirm before modifying each file", 
                       variable=self.confirm_each).pack(side=tk.LEFT)
        
        # Buttons
        button_frame = ttk.Frame(self.frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="Preview Affected Files", 
                  command=self._preview_files).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="Process Files", 
                  command=self._process_files).pack(side=tk.LEFT)
        
        # Status area
        status_label = ttk.Label(self.frame, text="Status:")
        status_label.pack(anchor=tk.W, pady=(10, 5))
        
        self.status_text = scrolledtext.ScrolledText(self.frame, height=10, width=80, state=tk.DISABLED)
        self.status_text.configure(bg=DARK_ENTRY, fg=DARK_FG, insertbackground=DARK_FG)
        self.status_text.pack(fill=tk.BOTH, expand=True)
    
    def _select_folder(self):
        folder = filedialog.askdirectory(title="Select Project Folder with .mpr Files")
        if folder:
            self.project_folder.set(folder)
            self._log_status(f"Selected folder: {folder}")
    
    def _log_status(self, message):
        if self.status_text is None:
            # Fallback when UI not yet built
            print(message)
            return
        self.status_text.config(state=tk.NORMAL)
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.status_text.config(state=tk.DISABLED)
    
    def _get_mpr_files(self, folder: str) -> list[Path]:
        return list(Path(folder).rglob("*.mpr"))
    
    def _file_contains_block(self, file_path: Path) -> bool:
        try:
            text = file_path.read_text(encoding="utf-8")
            return MPR_TEXT_TO_REMOVE in text
        except Exception as e:
            self._log_status(f"Error reading {file_path.name}: {e}")
            return False
    
    def _preview_files(self):
        # Flow: validate folder -> gather .mpr -> filter containing block -> show list
        folder = self.project_folder.get()
        if not folder:
            messagebox.showwarning("No Folder", "Please select a project folder first.")
            return
        
        if not Path(folder).exists():
            messagebox.showerror("Invalid Folder", "Selected folder does not exist.")
            return
        
        self._log_status("Scanning for .mpr files...")
        mpr_files = self._get_mpr_files(folder)
        
        if not mpr_files:
            messagebox.showinfo("No Files", "No .mpr files found in the selected folder.")
            self._log_status("No .mpr files found.")
            return
        
        self._log_status(f"Found {len(mpr_files)} .mpr file(s). Checking for text block...")
        
        affected = [f for f in mpr_files if self._file_contains_block(f)]
        
        if not affected:
            messagebox.showinfo("No Matches", "No files contain the specified text block.")
            self._log_status("No files contain the text block.")
            return
        
        preview_list = "\n".join(f.name for f in affected[:20])
        if len(affected) > 20:
            preview_list += f"\n... and {len(affected) - 20} more"
        
        self._log_status(f"Found {len(affected)} file(s) that contain the text block.")
        
        messagebox.showinfo(
            "Preview",
            f"{len(affected)} file(s) will be modified.\n\nFirst 20 files:\n{preview_list}"
        )
    
    def _create_backup(self, file_path: Path):
        backup_path = file_path.with_suffix(file_path.suffix + ".bak")
        if not backup_path.exists():
            try:
                backup_path.write_text(file_path.read_text(encoding="utf-8"), encoding="utf-8")
                return True
            except Exception as e:
                self._log_status(f"Error creating backup for {file_path.name}: {e}")
                return False
        return True
    
    def _process_files(self):
        # Flow: validate folder -> gather affected -> confirm -> backup -> replace block
        folder = self.project_folder.get()
        if not folder:
            messagebox.showwarning("No Folder", "Please select a project folder first.")
            return
        
        if not Path(folder).exists():
            messagebox.showerror("Invalid Folder", "Selected folder does not exist.")
            return
        
        self._log_status("Starting file processing...")
        mpr_files = self._get_mpr_files(folder)
        
        if not mpr_files:
            messagebox.showinfo("No Files", "No .mpr files found.")
            return
        
        affected = [f for f in mpr_files if self._file_contains_block(f)]
        
        if not affected:
            messagebox.showinfo("No Matches", "No files contain the specified text block.")
            return
        
        # Confirm before processing
        proceed = messagebox.askyesno(
            "Confirm Processing",
            f"{len(affected)} file(s) will be modified.\n\nProceed?"
        )
        
        if not proceed:
            self._log_status("Processing cancelled by user.")
            return
        
        modified_count = 0
        skipped_count = 0
        confirm_all = False

        # If user wants confirmations, offer a one-click confirm-all option up front
        if self.confirm_each.get():
            confirm_all = messagebox.askyesno(
                "Confirm All?",
                f"{len(affected)} file(s) will be modified.\n\n"
                "Click 'Yes' to confirm all files at once.\n"
                "Click 'No' to confirm each file individually."
            )
        
        for file_path in affected:
            try:
                if self.confirm_each.get() and not confirm_all:
                    answer = messagebox.askyesno(
                        "Confirm Modification",
                        f"Modify this file?\n\n{file_path.name}"
                    )
                    if not answer:
                        skipped_count += 1
                        self._log_status(f"Skipped: {file_path.name}")
                        continue
                
                # Create backup
                self._create_backup(file_path)
                
                # Remove text block
                original_text = file_path.read_text(encoding="utf-8")
                updated_text = original_text.replace(MPR_TEXT_TO_REMOVE, "")
                
                if updated_text != original_text:
                    file_path.write_text(updated_text, encoding="utf-8")
                    modified_count += 1
                    self._log_status(f"Modified: {file_path.name}")
                else:
                    self._log_status(f"No changes needed: {file_path.name}")
                    
            except Exception as e:
                self._log_status(f"Error processing {file_path.name}: {e}")
        
        message = f"Processing complete.\nModified: {modified_count} file(s)"
        if skipped_count > 0:
            message += f"\nSkipped: {skipped_count} file(s)"
        
        messagebox.showinfo("Complete", message)
        self._log_status(message)


class CutlistGeneratorTab:
    """Tab 2: Cutlist Generator"""
    
    def __init__(self, parent):
        self.frame = ttk.Frame(parent, padding="10")
        self.frame.pack(fill=tk.BOTH, expand=True)
        
        self.convention_file = tk.StringVar()
        self.cutlist_file = tk.StringVar()
        self.convention_df = None
        # Paths/config
        self.config_path = get_config_path()
        self.base_dir = get_app_base_dir()
        self.config_data = self._load_config()
        self.convention_json_path = self._get_configured_convention_path()
        self.edge_dir_path = self._get_configured_edge_dir()
        self.convention_path_var = tk.StringVar(value=str(self.convention_json_path))
        self.edge_dir_var = tk.StringVar(value=str(self.edge_dir_path))
        self.cutlist_df = None
        self.status_text = None
        self.tool_diameter = tk.DoubleVar(value=10.0)
        self.remove_macro_124 = tk.BooleanVar(value=False)
        
        self._create_widgets()
        self._update_path_entries()

    # ---------------- Config helpers ----------------
    def _load_config(self) -> dict:
        if not self.config_path.exists():
            return {}
        try:
            return json.loads(self.config_path.read_text(encoding="utf-8"))
        except Exception as e:
            self._log_status(f"Failed to load config ({self.config_path}): {e}")
            return {}

    def _save_config(self):
        data = {
            "convention_json": str(self.convention_json_path),
            "edge_dir": str(self.edge_dir_path),
        }
        try:
            self.config_path.write_text(json.dumps(data, indent=2), encoding="utf-8")
            self.config_data = data
            self._log_status(f"Saved defaults to {self.config_path}")
        except Exception as e:
            messagebox.showerror("Save Error", f"Could not save defaults:\n{e}")

    def _get_configured_convention_path(self) -> Path:
        candidate = self.config_data.get("convention_json")
        if candidate:
            return Path(candidate)
        return self.base_dir / "Polyboard_convention.json"

    def _get_configured_edge_dir(self) -> Path:
        candidate = self.config_data.get("edge_dir")
        if candidate:
            return Path(candidate)
        return self.base_dir / "Edge_Diagram_Ref"

    def _update_path_entries(self):
        self.convention_path_var.set(str(self.convention_json_path))
        if self.convention_path_label:
            self.convention_path_label.config(state="normal")
            self.convention_path_label.delete(0, tk.END)
            self.convention_path_label.insert(0, str(self.convention_json_path))
            self.convention_path_label.config(state="readonly")
        self.edge_dir_var.set(str(self.edge_dir_path))
    
    def _create_widgets(self):
        # Title
        title = ttk.Label(self.frame, text="Cutlist Generator", font=("Arial", 14, "bold"))
        title.pack(pady=(0, 10))
        
        # File selection frame
        file_frame = ttk.LabelFrame(self.frame, text="File Selection", padding="10")
        file_frame.pack(fill=tk.X, pady=5)
        
        # Convention info (read-only JSON path) + editor
        conv_frame = ttk.Frame(file_frame)
        conv_frame.pack(fill=tk.X, pady=5)
        ttk.Label(conv_frame, text="Convention JSON:").pack(side=tk.LEFT, padx=(0, 5))
        self.convention_path_label = ttk.Entry(conv_frame, textvariable=self.convention_path_var, width=60, state="readonly")
        self.convention_path_label.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
        ttk.Button(conv_frame, text="Browse", command=self._choose_convention_json).pack(side=tk.LEFT, padx=(5, 0))
        ttk.Button(conv_frame, text="Edit Convention", command=self._open_convention_editor).pack(side=tk.LEFT, padx=(5, 0))
        
        # Edge diagram folder
        edge_frame = ttk.Frame(file_frame)
        edge_frame.pack(fill=tk.X, pady=5)
        ttk.Label(edge_frame, text="Edge Diagram Folder:").pack(side=tk.LEFT, padx=(0, 5))
        self.edge_dir_entry = ttk.Entry(edge_frame, textvariable=self.edge_dir_var, width=60, state="readonly")
        self.edge_dir_entry.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
        ttk.Button(edge_frame, text="Browse", command=self._choose_edge_dir).pack(side=tk.LEFT)
        
        # Save defaults
        save_frame = ttk.Frame(file_frame)
        save_frame.pack(fill=tk.X, pady=2)
        ttk.Button(save_frame, text="Save defaults (paths above)", command=self._save_defaults).pack(side=tk.LEFT, padx=(0, 5))
        
        # Cutlist file
        cutlist_frame = ttk.Frame(file_frame)
        cutlist_frame.pack(fill=tk.X, pady=5)
        ttk.Label(cutlist_frame, text="Project Cutlist CSV:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(cutlist_frame, textvariable=self.cutlist_file, width=50).pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
        ttk.Button(cutlist_frame, text="Browse", command=self._select_cutlist_file).pack(side=tk.LEFT)
        
        # Buttons
        button_frame = ttk.Frame(self.frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="Load & Preview", 
                  command=self._load_and_preview).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="Export Final Cutlist", 
                  command=self._export_cutlist).pack(side=tk.LEFT)
        ttk.Label(button_frame, text="Tool Ø (mm):").pack(side=tk.LEFT, padx=(10, 4))
        ttk.Spinbox(button_frame, from_=1, to=50, increment=0.5, textvariable=self.tool_diameter, width=6).pack(side=tk.LEFT)
        ttk.Checkbutton(button_frame, text="Remove macro 124", variable=self.remove_macro_124).pack(side=tk.LEFT, padx=(10, 0))

        # Preview area (table with horizontal + vertical scroll)
        preview_label = ttk.Label(self.frame, text="Preview (first rows):")
        preview_label.pack(anchor=tk.W, pady=(10, 5))

        self.preview_frame = ttk.Frame(self.frame)
        self.preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        self.preview_tree = ttk.Treeview(self.preview_frame, columns=(), show="headings", height=12)
        vsb = ttk.Scrollbar(self.preview_frame, orient="vertical", command=self.preview_tree.yview)
        hsb = ttk.Scrollbar(self.preview_frame, orient="horizontal", command=self.preview_tree.xview)
        self.preview_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.preview_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.preview_frame.rowconfigure(0, weight=1)
        self.preview_frame.columnconfigure(0, weight=1)
        
        # Status area
        status_label = ttk.Label(self.frame, text="Status:")
        status_label.pack(anchor=tk.W, pady=(5, 5))
        
        self.status_text = scrolledtext.ScrolledText(self.frame, height=8, width=100, state=tk.DISABLED)
        self.status_text.configure(bg=DARK_ENTRY, fg=DARK_FG, insertbackground=DARK_FG)
        self.status_text.pack(fill=tk.BOTH, expand=True)
    
    def _select_convention_file(self):
        file = filedialog.askopenfilename(
            title="Select Convention File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file:
            self.convention_file.set(file)
            self._log_status(f"Selected convention file: {file}")

    def _choose_convention_json(self):
        file = filedialog.askopenfilename(
            title="Select Convention JSON",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if file:
            self.convention_json_path = Path(file)
            self._update_path_entries()
            self._log_status(f"Selected convention JSON: {file}")
    
    def _select_cutlist_file(self):
        file = filedialog.askopenfilename(
            title="Select Project Cutlist CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if file:
            self.cutlist_file.set(file)
            if self.convention_path_label:
                self.convention_path_label.config(state="normal")
                self.convention_path_label.delete(0, tk.END)
                self.convention_path_label.insert(0, str(self.convention_json_path))
                self.convention_path_label.config(state="readonly")
            self._log_status(f"Selected cutlist file: {file}")

    def _choose_edge_dir(self):
        folder = filedialog.askdirectory(
            title="Select Edge_Diagram_Ref Folder"
        )
        if folder:
            self.edge_dir_path = Path(folder)
            self._update_path_entries()
            self._log_status(f"Selected edge diagram folder: {folder}")

    def _save_defaults(self):
        self._save_config()
        messagebox.showinfo("Defaults Saved", f"Saved to {self.config_path}")
    
    def _log_status(self, message):
        self.status_text.config(state=tk.NORMAL)
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.status_text.config(state=tk.DISABLED)
    
    def _load_convention_file(self):
        """Load and validate convention file"""
        # Always load from JSON; create empty if missing
        json_path = self._get_convention_json_path()
        if self.convention_path_label:
            self.convention_path_label.config(state="normal")
            self.convention_path_label.delete(0, tk.END)
            self.convention_path_label.insert(0, str(json_path))
            self.convention_path_label.config(state="readonly")

        if json_path.exists():
            try:
                data = json.loads(json_path.read_text(encoding="utf-8"))
                df = pd.DataFrame(data)
                missing = [col for col in CONVENTION_COLUMNS if col not in df.columns]
                if missing:
                    # Auto-fill missing columns to avoid hard failure
                    for col in missing:
                        df[col] = ""
                    self._log_status(f"Convention JSON missing columns {missing}; filled with empty values.")
                df = df[[c for c in CONVENTION_COLUMNS]]
                self._log_status(f"Loaded convention JSON: {json_path}")
                self.convention_df = df
                return df
            except Exception as e:
                self._log_status(f"Failed to load convention JSON: {e}")
                raise

        # Create empty convention if none exists
        self._log_status("Convention JSON not found; creating empty convention set.")
        df = pd.DataFrame(columns=CONVENTION_COLUMNS)
        self.convention_df = df
        # Save empty JSON so editor/loader share path
        try:
            json_path.parent.mkdir(parents=True, exist_ok=True)
            json_path.write_text(df.to_json(orient="records", force_ascii=False, indent=2), encoding="utf-8")
        except Exception as e:
            self._log_status(f"Failed to write empty convention JSON: {e}")
        return df
    
    def _load_cutlist_file(self):
        """Load and validate cutlist CSV file"""
        file_path = self.cutlist_file.get()
        if not file_path:
            raise ValueError("No cutlist file selected.")
        
        if not Path(file_path).exists():
            raise FileNotFoundError(f"Cutlist file not found: {file_path}")
        
        try:
            # Read CSV with semicolon separator; files have no headers so we assign them
            df = pd.read_csv(
                file_path,
                sep=';',
                encoding='utf-8',
                header=None,
                names=CUTLIST_COLUMNS,
            )
            self._log_status(f"Loaded cutlist file: {len(df)} rows")
            return df
        except Exception as e:
            raise Exception(f"Error loading cutlist file: {e}")
    
    def _match_component(self, reference: str, convention_df: pd.DataFrame) -> str:
        """Match Reference to Component using specific rules"""
        reference_upper = reference.upper()
        
        # Priority 1: Check for "L Side Drawer" or "R Side Drawer"
        if "L SIDE DRAWER" in reference_upper or "R SIDE DRAWER" in reference_upper:
            for component in convention_df['Component']:
                if component == "Drawers Side":
                    return component
        
        # Priority 2: Check for "Drawer" (without "Side")
        if "DRAWER" in reference_upper and "SIDE DRAWER" not in reference_upper:
            for component in convention_df['Component']:
                if component == "Drawers (Frontage)":
                    return component

        # Priority 3: Double doors
        if "DOOR" in reference_upper and "DOUBLE" in reference_upper:
            for component in convention_df['Component']:
                if component == "Doors (Double)":
                    return component

        # Priority 3: Single Door (Std) open vs fitting side
        if "DOOR" in reference_upper and "SINGLE" in reference_upper:
            target = "Single Doors (Open Side)" if "OPEN" in reference_upper else "Single Doors (Fitting Side)"
            for component in convention_df['Component']:
                if component == target:
                    return component
        
        # Priority 4: Try exact match
        for component in convention_df['Component']:
            if reference == component:
                return component
        
        # Priority 5: Try partial match (remove brackets and numbers)
        # Clean reference: remove [1], [2], etc.
        cleaned_ref = re.sub(r'\s*\[\d+\]', '', reference).strip()
        for component in convention_df['Component']:
            if cleaned_ref == component:
                return component
            # Also try if cleaned reference contains component
            if component in cleaned_ref or cleaned_ref in component:
                return component
        
        # Priority 6: Try case-insensitive partial match
        for component in convention_df['Component']:
            if component.upper() in reference_upper or reference_upper in component.upper():
                return component
        
        return None
    
    def _count_edges(self, row: pd.Series, edge_columns: list) -> tuple:
        """Count edge banding and determine if 2 edges are adjacent or opposite"""
        edges = []
        for col in edge_columns:
            value = str(row[col]) if pd.notna(row[col]) else ""
            if value.strip() and value.strip() != "nan":
                edges.append(True)
            else:
                edges.append(False)
        
        edge_count = sum(edges)
        
        # Determine adjacency for 2-edge case
        is_adjacent = None
        if edge_count == 2:
            # Check if opposite: (Right + Left) or (Bottom + Top)
            # Assuming order: Right, Left, Bottom, Top
            if len(edges) >= 4:
                right_left = edges[0] and edges[1]  # Right and Left
                bottom_top = edges[2] and edges[3]   # Bottom and Top
                is_adjacent = not (right_left or bottom_top)
            else:
                is_adjacent = True  # Default if we can't determine
        
        return edge_count, is_adjacent
    
    def _get_edge_code(self, component: str, edge_count: int, is_adjacent: bool, convention_df: pd.DataFrame) -> str:
        """Get edge code from convention file based on component, edge count, and adjacency"""
        if component is None:
            return ""
        
        component_row = convention_df[convention_df['Component'] == component]
        if component_row.empty:
            return ""
        
        if edge_count == 0:
            col = 'Edge_0'
        elif edge_count == 1:
            col = 'Edge_1'
        elif edge_count == 2:
            col = 'Edge_2_connect' if is_adjacent else 'Edge_2_no_connect'
        elif edge_count == 3:
            col = 'Edge_3'
        elif edge_count == 4:
            col = 'Edge_4'
        else:
            return ""
        
        value = component_row[col].iloc[0]
        if pd.isna(value):
            return ""
        
        return str(value)
    
    def _convert_grain_direction(self, value):
        """Convert grain direction: 0→N, 1→Y, 2→X"""
        if pd.isna(value):
            return ""
        
        value_str = str(value).strip()
        if value_str == "0":
            return "N"
        elif value_str == "1":
            return "Y"
        elif value_str == "2":
            return "X"
        else:
            return value_str  # Return as-is if not 0, 1, or 2
    
    def _process_cutlist(self):
        """Process cutlist with all conversions (grain, component match, edge code, face name)"""
        if self.convention_df is None or self.cutlist_df is None:
            raise ValueError("Please load files first.")
        
        df = self.cutlist_df.copy()
        
        # Use predefined columns (since input has no headers)
        reference_col = "Reference"
        grain_col = "Grain_Direction"
        edging_col = "Edging_Diagram"
        face_col = "Face"
        edge_columns = ["Right_Edge", "Left_Edge", "Bottom_Edge", "Top_Edge"]
        
        unmatched_components = []
        
        # Process each row
        for idx, row in df.iterrows():
            reference = str(row[reference_col]) if pd.notna(row[reference_col]) else ""
            
            # 1. Convert Grain Direction
            if grain_col and grain_col in df.columns:
                df.at[idx, grain_col] = self._convert_grain_direction(row[grain_col])
            
            # 2. Match component
            component = self._match_component(reference, self.convention_df)
            
            if component is None:
                unmatched_components.append(reference)
                continue
            
            # 3. Count edges and determine adjacency
            edge_count, is_adjacent = self._count_edges(row, edge_columns)
            
            # 4. Get edge code
            edge_code = self._get_edge_code(component, edge_count, is_adjacent, self.convention_df)
            
            # 5. Update Edging_Diagram
            if edging_col and edging_col in df.columns:
                df.at[idx, edging_col] = edge_code
            
            # 6. Add Face Name
            if face_col and face_col in df.columns:
                face_value = row[face_col]
                if pd.notna(face_value):
                    face_num = str(face_value).strip()
                    component_row = self.convention_df[self.convention_df['Component'] == component]
                    if not component_row.empty:
                        if face_num == "1":
                            face_name = component_row['Face_1'].iloc[0]
                        elif face_num == "2":
                            face_name = component_row['Face_2'].iloc[0]
                        else:
                            face_name = ""
                        
                        # Add Face Name column if it doesn't exist
                        if "Face Name" not in df.columns:
                            df.insert(len(df.columns), "Face Name", "")
                        
                        df.at[idx, "Face Name"] = face_name if pd.notna(face_name) else ""
        
        if unmatched_components:
            self._log_status(f"Warning: {len(unmatched_components)} unmatched components: {set(unmatched_components)}")
        
        return df

    # ---------------- MPR validation & process summary helpers ----------------
    def _resolve_project_folder(self) -> Path:
        if self.cutlist_file.get():
            return Path(self.cutlist_file.get()).parent
        return Path.cwd()

    def _get_convention_json_path(self) -> Path:
        path = Path(self.convention_json_path)
        if not path.exists():
            fallback = self.base_dir / "Polyboard_convention.json"
            self._log_status(f"Convention path not found, falling back to {fallback}")
            self.convention_json_path = fallback
            self._update_path_entries()
            return fallback
        return path

    def _generate_deterministic_id(self, *texts) -> str:
        """Deterministic UUID (uuid5) from sorted text inputs."""
        sorted_texts = sorted(map(str, texts))
        combined = "|".join(sorted_texts)
        return str(uuid.uuid5(uuid.NAMESPACE_DNS, combined))

    def _collect_mpr_names(self, df: pd.DataFrame) -> list[str]:
        mpr_col = "Tooling_File_First_Face"
        if mpr_col not in df.columns:
            return []
        return [
            str(v).strip()
            for v in df[mpr_col].tolist()
            if pd.notna(v) and str(v).strip()
        ]

    def _locate_mpr_files(self, mpr_names: list[str], project_folder: Path) -> tuple[dict, list[str]]:
        found_map: dict[str, Path] = {}
        missing: list[str] = []
        for name in sorted(set(mpr_names)):
            direct = project_folder / name
            if direct.exists():
                found_map[name] = direct
                continue
            matches = list(project_folder.rglob(name))
            if matches:
                found_map[name] = matches[0]
            else:
                missing.append(name)
        return found_map, missing

    def _strip_macro_124(self, file_path: Path):
        """Remove all macro 124 blocks from an MPR file; keep .bak if not present."""
        try:
            backup = file_path.with_suffix(file_path.suffix + ".bak")
            if not backup.exists():
                backup.write_text(file_path.read_text(encoding="utf-8"), encoding="utf-8")

            text = file_path.read_text(encoding="utf-8")
            pattern = re.compile(r'(?ms)^\\s*<\\s*124\\s*\\\\.*?(?=^\\s*<\\s*\\d+\\s*\\\\|\\Z)')
            cleaned = pattern.sub("", text)
            if cleaned != text:
                file_path.write_text(cleaned, encoding="utf-8")
                self._log_status(f"Removed macro 124 from {file_path.name}")
        except Exception as e:
            self._log_status(f"Failed to strip macro 124 from {file_path}: {e}")

    def _remove_component_block(self, text: str) -> tuple[str, bool]:
        cleaned = text.replace(MPR_TEXT_TO_REMOVE, "")
        return cleaned, cleaned != text

    def _get_param(self, block: str, key: str) -> str:
        # More permissive: search anywhere in the block, case-insensitive
        pattern = re.compile(rf'(?i){re.escape(key)}\s*=\s*(?:"([^"]*)"|([^\s\\\r\n]+))')
        m = pattern.search(block)
        if not m:
            return ""
        return (m.group(1) if m.group(1) is not None else m.group(2)).strip()

    def _parse_macro_100_dims(self, text: str) -> tuple[float, float]:
        """Extract LA/BR from macro 100; tolerant to whitespace/backslashes."""
        la = br = 0.0
        match = re.search(r'(?ms)^\s*<\s*100\b.*?(?=^\s*<\s*\d+\b|\Z)', text)
        if match:
            block = match.group(0)
            la_val = self._get_param(block, "LA")
            br_val = self._get_param(block, "BR")
            try:
                la = float(la_val)
            except ValueError:
                la = 0.0
            try:
                br = float(br_val)
            except ValueError:
                br = 0.0
        return la, br

    def _convert_109_to_151(self, block: str, dims: tuple[float, float], tool_diam: float):
        la_100, br_100 = dims
        xa = self._get_param(block, "XA")
        ya = self._get_param(block, "YA")
        xe = self._get_param(block, "XE")
        ye = self._get_param(block, "YE")
        nb = self._get_param(block, "NB")
        ti = self._get_param(block, "TI")
        tval = self._get_param(block, "T_")
        rk = self._get_param(block, "RK")

        try:
            xa_f = float(xa); ya_f = float(ya); xe_f = float(xe); ye_f = float(ye)
        except ValueError:
            return None

        dx = abs(xa_f - xe_f)
        dy = abs(ya_f - ye_f)
        ddx = xe_f - xa_f
        ddy = ye_f - ya_f
        along_x = dx != 0
        groove_len = dx if along_x else dy

        try:
            nb_f = float(nb) if nb else 0.0
        except ValueError:
            nb_f = 0.0

        rk_norm = rk.upper() if isinstance(rk, str) else "NOWRK"

        if along_x:
            xa151 = la_100 / 2 if la_100 else xa_f
            ya151 = ya_f
            if rk_norm == "WRKL":
                if ddx > 0:
                    ya151 = ya_f + (nb_f / 2)
                else:
                    ya151 = ya_f - (nb_f / 2)
            elif rk_norm == "WRKR":
                if ddx > 0:
                    ya151 = ya_f - (nb_f / 2)
                else:
                    ya151 = ya_f + (nb_f / 2)
            la151 = groove_len + tool_diam
            br151 = nb_f
        else:
            xa151 = xa_f
            if rk_norm == "WRKL":
                if ddy > 0:
                    xa151 = xa_f - (nb_f / 2)
                else:
                    xa151 = xa_f + (nb_f / 2)
            elif rk_norm == "WRKR":
                if ddy > 0:
                    xa151 = xa_f + (nb_f / 2)
                else:
                    xa151 = xa_f - (nb_f / 2)
            ya151 = br_100 / 2 if br_100 else ya_f
            la151 = nb_f
            br151 = groove_len + tool_diam

        ti151 = ti

        new_block = [
            '<151 \\UflurTasche\\',
            f'XA="{xa151}"',
            f'YA="{ya151}"',
            f'LA="{la151}"',
            f'BR="{br151}"',
            f'TI="{ti151}"',
            'RD="0"',
            'WI="0"',
            'ZT="0"',
            'XY="80"',
            'AB="30"',
            'AM="1"',
            'DS="0"',
            'T_="3"',
            'KO="00"',
            ''
        ]
        return "\n".join(new_block), ("X" if along_x else "Y"), groove_len

    def _transform_mpr(self, path: Path, tool_diam: float, remove_macro_124: bool) -> dict:
        """Return transformed text and actions without writing."""
        actions = {
            "path": path,
            "removed_component": False,
            "removed_124": False,
            "remove_124_requested": remove_macro_124,
            "remove_124_skipped": False,
            "has_macro_124": False,
            "converted_109_to_151": [],
            "new_text": None,
            "changed": False,
            "la_100": 0.0,
            "br_100": 0.0,
        }
        try:
            text = path.read_text(encoding="utf-8")
        except Exception as e:
            actions["error"] = f"Read error: {e}"
            return actions

        la_100, br_100 = self._parse_macro_100_dims(text)
        actions["la_100"] = la_100
        actions["br_100"] = br_100

        # Remove component block
        text, removed_comp = self._remove_component_block(text)
        actions["removed_139_InvalidMacro"] = removed_comp
        actions["removed_component"] = removed_comp

        # Detect macro 124 presence (before any optional removal)
        pattern124 = re.compile(r'(?ms)^\s*<\s*124\b.*?(?=^\s*<\s*\d+\b|\Z)')
        actions["has_macro_124"] = bool(pattern124.search(text))

        # Remove 124 (optional)
        if remove_macro_124:
            cleaned_124 = pattern124.sub("", text)
            if cleaned_124 != text:
                actions["removed_124"] = True
                text = cleaned_124
        else:
            actions["remove_124_skipped"] = True

        # Process blocks
        block_re = re.compile(r'(?ms)^\s*<\s*(\d+)\b.*?(?=^\s*<\s*\d+\b|\Z)')
        matches = list(block_re.finditer(text))
        out_blocks = []
        for m in matches:
            block = m.group(0)
            mid = m.group(1)
            if mid == "124" and remove_macro_124:
                continue
            if mid == "109":
                t_val = self._get_param(block, "T_")
                t_clean = t_val.replace('"', "").replace("!", "").strip()
                if t_clean.endswith("xxxxx2"):
                    self._log_status(f"Tool in 2 face [({la_100}, {br_100}), {tool_diam}] for {path.name}")
                    conv = self._convert_109_to_151(block, (la_100, br_100), tool_diam)
                    if conv:
                        self._log_status(f"Converted 109 to 151 for {path.name} successfully")
                        new151, axis, glen = conv
                        actions["converted_109_to_151"].append(f"{path.name} axis={axis} L={glen}")
                        out_blocks.append(new151)
                        actions["changed"] = True
                        continue
                else:
                    self._log_status(f"Skipped 109 (No T_ or T_ not ending with 2 meaning Face_1 Grv) in {path.name}")
            out_blocks.append(block)

        if not out_blocks:
            actions["new_text"] = text
            return actions

        new_text = "\n".join(out_blocks)
        if new_text != text:
            actions["changed"] = True
        actions["new_text"] = new_text
        return actions

    def _show_mpr_changes_dialog(self, actions: list[dict]) -> bool:
        dialog = tk.Toplevel(self.frame)
        dialog.title("Confirm MPR Changes")
        dialog.geometry("900x400")
        dialog.transient(self.frame.winfo_toplevel())
        dialog.grab_set()
        apply_dark_theme(dialog)

        cols = ["File", "LA_100", "BR_100", "removed_139_InvalidMacro", "Angle_Bevel_Operation_Remove", "Convert_To_Milling_Face2_Grv", "Error"]
        tree = ttk.Treeview(dialog, columns=cols, show="headings", height=12)
        vsb = ttk.Scrollbar(dialog, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(dialog, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=160, anchor="w")

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        dialog.rowconfigure(0, weight=1)
        dialog.columnconfigure(0, weight=1)

        for act in actions:
            fname = Path(act["path"]).name if act.get("path") else ""
            la100 = act.get("la_100", "")
            br100 = act.get("br_100", "")
            comp = "yes" if act.get("removed_139_InvalidMacro") else ""
            m124 = ""
            if act.get("has_macro_124"):
                if act.get("removed_124"):
                    m124 = "removed"
                elif act.get("remove_124_requested") is False:
                    m124 = "kept (toggle off)"
                else:
                    m124 = "present"
            conv = ", ".join(act.get("converted_109_to_151", []))
            err = act.get("error", "")
            tree.insert("", "end", values=[fname, la100, br100, comp, m124, conv, err])

        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=2, column=0, sticky="e", pady=8, padx=8)

        result = {"proceed": False}

        def on_ok():
            result["proceed"] = True
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        ttk.Button(btn_frame, text="Proceed", command=on_ok).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_frame, text="Cancel", command=on_cancel).pack(side=tk.LEFT, padx=4)

        dialog.wait_window()
        return result["proceed"]

    def _summarize_mpr_processes(self, found_map: dict[str, Path]) -> dict[str, dict]:
        summary_cache: dict[str, dict] = {}
        for name, path in found_map.items():
            try:
                result = map_and_count_mpr_processes(str(path))
                mapped = result.get("mapped_process_counts", {})
                vert_sig = result.get("bohrvert_signature_counts", {})
                horiz_sig = result.get("bohrhoriz_signature_counts", {})
                angle_lengths = result.get("angle124_lengths", [])
                groove109_lengths = result.get("groove109_lengths", [])

                # Summary parts: include totals for all macros (including 102/103)
                parts = []
                for mid, info in sorted(mapped.items()):
                    desc = info.get("description", "")
                    cnt = info.get("count", 0)
                    if cnt <= 0 or desc == "Unknown/Unmapped macro ID":
                        continue
                    if mid == 124 and angle_lengths:
                        length_vals = [str(l).split("_")[0] for l in angle_lengths]
                        length_str = ",".join(length_vals)
                        parts.append(f"{desc}:{cnt} [L={length_str}]")
                    elif mid == 109 and groove109_lengths:
                        length_vals = [str(l).split("_")[0] for l in groove109_lengths]
                        length_str = ",".join(length_vals)
                        parts.append(f"{desc}:{cnt} [L={length_str}]")
                    else:
                        parts.append(f"{desc}:{cnt}")

                # Detail strings for 102/103 signatures (JSON-like)
                vert_detail = ""
                if vert_sig:
                    vert_detail = "@".join([f"{k}:{v}" for k, v in sorted(vert_sig.items())])
                horiz_detail = ""
                if horiz_sig:
                    horiz_detail = "@".join([f"{k}:{v}" for k, v in sorted(horiz_sig.items())])

                summary_cache[name] = {
                    "summary": "@".join(parts),
                    "vert": vert_detail,
                    "horiz": horiz_detail,
                    "angle_len": "@".join([f"{l}" for l in angle_lengths]) if angle_lengths else "",
                    "groove109_len": "@".join([f"{l}" for l in groove109_lengths]) if groove109_lengths else "",
                }
            except Exception as e:
                summary_cache[name] = {
                    "summary": f"ERROR: {e}",
                    "vert": "",
                    "horiz": "",
                    "angle_len": "",
                    "groove109_len": "",
                }
                self._log_status(f"Process summary error for {name}: {e}")
        return summary_cache

    def _validate_and_annotate_mprs(self, processed_df: pd.DataFrame) -> pd.DataFrame | None:
        """Ensure referenced MPRs exist; if all found, append process summary column."""
        project_folder = self._resolve_project_folder()
        mpr_names = self._collect_mpr_names(processed_df)
        found_map, missing = self._locate_mpr_files(mpr_names, project_folder)

        if missing:
            miss_str = "\n".join(sorted(set(missing)))
            messagebox.showwarning(
                "Missing MPR Files",
                f"The following MPR files are missing in the project folder:\n\n{miss_str}"
            )
            self._log_status(f"Operation aborted. Missing MPR files: {set(missing)}")
            return None

        summary_cache = self._summarize_mpr_processes(found_map)

        if "Tooling_File_First_Face" in processed_df.columns:
            processed_df["MPR_Process_Summary"] = processed_df["Tooling_File_First_Face"].apply(
                lambda v: summary_cache.get(str(v).strip(), {}).get("summary", "") if pd.notna(v) else ""
            )
            processed_df["Vertical_Drill_Detail"] = processed_df["Tooling_File_First_Face"].apply(
                lambda v: summary_cache.get(str(v).strip(), {}).get("vert", "") if pd.notna(v) else ""
            )
            processed_df["Horizontal_Drill_Detail"] = processed_df["Tooling_File_First_Face"].apply(
                lambda v: summary_cache.get(str(v).strip(), {}).get("horiz", "") if pd.notna(v) else ""
            )
            processed_df["Angle_Groove_Length"] = processed_df["Tooling_File_First_Face"].apply(
                lambda v: summary_cache.get(str(v).strip(), {}).get("angle_len", "") if pd.notna(v) else ""
            )
            processed_df["Saw_Groove_Length"] = processed_df["Tooling_File_First_Face"].apply(
                lambda v: summary_cache.get(str(v).strip(), {}).get("groove109_len", "") if pd.notna(v) else ""
            )
            # Edge band count
            edge_cols = ["Right_Edge", "Left_Edge", "Bottom_Edge", "Top_Edge"]
            def _count_edges(row):
                count = 0
                for c in edge_cols:
                    val = row.get(c, "")
                    if pd.notna(val) and str(val).strip() and str(val).strip() != "nan":
                        count += 1
                return count
            processed_df["Edge_Band_Count"] = processed_df.apply(_count_edges, axis=1)

            # Unique deterministic ID
            def _make_uid(row):
                ref = row.get("Reference", "")
                proj = row.get("Project", "")
                cab = row.get("Cabinet", "")
                cln = row.get("Cutting_List_Number", "")
                return self._generate_deterministic_id(proj, cab, ref, cln)
            processed_df["Unique_ID"] = processed_df.apply(_make_uid, axis=1)

        return processed_df

    def _on_convention_saved(self, df: pd.DataFrame):
        self.convention_df = df
        self._log_status(f"Convention updated in memory ({len(df)} rows).")
        json_path = self._get_convention_json_path()
        try:
            json_path.parent.mkdir(parents=True, exist_ok=True)
            json_path.write_text(df.to_json(orient="records", force_ascii=False, indent=2), encoding="utf-8")
            self._log_status(f"Convention JSON saved to {json_path}")
        except Exception as e:
            self._log_status(f"Failed to save convention JSON: {e}")
        # Refresh label/path display
        if self.convention_path_label:
            self.convention_path_label.config(state="normal")
            self.convention_path_label.delete(0, tk.END)
            self.convention_path_label.insert(0, str(json_path))
            self.convention_path_label.config(state="readonly")
        # Immediately reload to ensure downstream uses latest data
        try:
            self.convention_df = self._load_convention_file()
        except Exception as e:
            self._log_status(f"Failed to reload convention after save: {e}")

    def _open_convention_editor(self):
        try:
            current_df = self._load_convention_file()
        except Exception as e:
            messagebox.showerror("Load Error", f"Unable to load convention:\n{e}")
            current_df = pd.DataFrame(columns=CONVENTION_COLUMNS)
        json_path = self._get_convention_json_path()
        ConventionEditorDialog(self.frame.winfo_toplevel(), current_df, json_path, self._on_convention_saved, edge_dir=self.edge_dir_path)

    def _populate_preview(self, df: pd.DataFrame, max_rows: int = 50):
        """Populate the preview treeview with first rows of df"""
        # Clear existing
        for col in self.preview_tree["columns"]:
            self.preview_tree.heading(col, text="")
            self.preview_tree.column(col, width=0)
        self.preview_tree.delete(*self.preview_tree.get_children())

        cols = list(df.columns)
        self.preview_tree["columns"] = cols
        for col in cols:
            self.preview_tree.heading(col, text=col)
            self.preview_tree.column(col, width=140, anchor="w")

        for _, row in df.head(max_rows).iterrows():
            values = [row.get(col, "") for col in cols]
            self.preview_tree.insert("", "end", values=values)

        if len(df) > max_rows:
            self._log_status(f"Preview shows first {max_rows} of {len(df)} rows.")
        else:
            self._log_status(f"Preview shows all {len(df)} rows.")
    
    def _load_and_preview(self):
        """Load files and show preview"""
        try:
            self._log_status("Loading files...")
            self.convention_df = self._load_convention_file()
            self.cutlist_df = self._load_cutlist_file()
            
            self._log_status("Processing cutlist...")
            processed_df = self._process_cutlist()

            # Validate MPRs and append process summary during preview
            processed_df = self._validate_and_annotate_mprs(processed_df)
            if processed_df is None:
                return
            
            # Show preview in table
            self._populate_preview(processed_df)
            self._log_status("Preview generated successfully.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error loading/processing files:\n{e}")
            self._log_status(f"Error: {e}")
    
    def _export_cutlist(self):
        """Export processed cutlist to CSV"""
        if self.convention_df is None or self.cutlist_df is None:
            messagebox.showwarning("No Data", "Please load and preview files first.")
            return
        
        try:
            processed_df = self._process_cutlist()
            processed_df = self._validate_and_annotate_mprs(processed_df)
            if processed_df is None:
                return

            # Preprocess MPR files: remove component block, strip 124, convert 109->151
            mpr_names = self._collect_mpr_names(processed_df)
            project_folder = self._resolve_project_folder()
            found_map, _ = self._locate_mpr_files(mpr_names, project_folder)

            actions = []
            for path in found_map.values():
                actions.append(self._transform_mpr(path, self.tool_diameter.get(), self.remove_macro_124.get()))

            # Build confirmation summary
            lines = []
            for act in actions:
                if act.get("error"):
                    lines.append(f"{act['path'].name}: ERROR {act['error']}")
                    continue
                flags = []
                if act.get("removed_139_InvalidMacro") or act.get("removed_component"):
                    flags.append("removed_component")
                if act.get("removed_124"):
                    flags.append("removed_124")
                elif act.get("remove_124_requested") is False:
                    flags.append("kept_124 (toggle off)")
                elif act.get("remove_124_requested") and not act.get("removed_124"):
                    flags.append("no_124_found")
                if act["converted_109_to_151"]:
                    flags.extend(act["converted_109_to_151"])
                if not flags:
                    flags.append("no changes")
                lines.append(f"{act['path'].name}: " + "; ".join(flags))
            preview_str = "\n".join(lines[:15])
            if len(lines) > 15:
                preview_str += f"\n... and {len(lines) - 15} more"

            # Interactive confirmation dialog
            if not self._show_mpr_changes_dialog(actions):
                self._log_status("Export cancelled by user before MPR modifications.")
                return

            # Apply changes
            for act in actions:
                if act.get("error") or not act.get("changed"):
                    continue
                try:
                    path = act["path"]
                    backup = path.with_suffix(path.suffix + ".bak")
                    if not backup.exists():
                        backup.write_text(path.read_text(encoding="utf-8"), encoding="utf-8")
                    path.write_text(act["new_text"], encoding="utf-8")
                    self._log_status(f"Applied changes to {path.name}")
                except Exception as e:
                    self._log_status(f"Failed to write {act.get('path')}: {e}")

            # Build default file name: To_Cutrite_[YYMMDD]_[HHMM]_[original cutlist filename].csv
            default_name = "To_Cutrite"
            if self.cutlist_file.get():
                original = Path(self.cutlist_file.get()).name
                today = datetime.now().strftime("%y%m%d")
                now = datetime.now().strftime("%H%M")
                default_name = f"To_Cutrite_{today}_{now}_{original}"
            else:
                default_name = "To_Cutrite.csv"
            
            # Ask for save location
            output_file = filedialog.asksaveasfilename(
                title="Save Final Cutlist",
                defaultextension=".csv",
                initialfile=default_name,
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
            )
            
            if not output_file:
                return
            
            # Export to CSV with semicolon separator
            processed_df.to_csv(output_file, sep=';', index=False, encoding='utf-8')
            
            messagebox.showinfo("Success", f"Cutlist exported successfully to:\n{output_file}")
            self._log_status(f"Exported to: {output_file}")
            self._populate_preview(processed_df)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error exporting cutlist:\n{e}")
            self._log_status(f"Export error: {e}")


class PolyboardProductionGUI:
    """Main application window"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Polyboard Production Tools")
        self.root.geometry("900x700")
        
        # Only Cutlist tab (Tab1 removed; cleanup runs during export)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.cutlist_tab = CutlistGeneratorTab(self.notebook)
        self.notebook.add(self.cutlist_tab.frame, text="Cutlist Generator")


def main():
    root = tk.Tk()
    apply_dark_theme(root)
    app = PolyboardProductionGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

