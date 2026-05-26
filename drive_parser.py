import csv
import queue
import threading
import tkinter as tk
from dataclasses import astuple, dataclass
from tkinter import filedialog, messagebox, ttk

import requests

# ── Domain ────────────────────────────────────────────────────────────────────

COLUMNS = [
    ("folder",   "Folder",       200),
    ("name",     "Name",         260),
    ("mime",     "MIME Type",    200),
    ("size",     "Size (bytes)", 110),
    ("modified", "Modified",      90),
    ("id",       "File ID",      180),
]

FOLDER_MIME = "application/vnd.google-apps.folder"
DRIVE_FILES_URL = "https://www.googleapis.com/drive/v3/files"


@dataclass
class DriveFile:
    folder: str
    name: str
    mime: str
    size: str
    modified: str
    id: str


# ── API (runs in background thread) ──────────────────────────────────────────

def scan_folder(token: str, folder_id: str, q: queue.Queue, stop: threading.Event) -> None:
    """Recursively scans a Drive folder. Posts to q: ('file', DriveFile),
    ('status', str), ('done', None), or ('error', str)."""
    try:
        _recurse(token, folder_id, q, stop)
        if not stop.is_set():
            q.put(("done", None))
    except Exception as exc:
        q.put(("error", str(exc)))


def _recurse(token: str, folder_id: str, q: queue.Queue, stop: threading.Event, path: str = "/") -> None:
    page_token = None
    headers = {"Authorization": f"Bearer {token}"}

    while not stop.is_set():
        params = {
            "q": f"'{folder_id}' in parents and trashed=false",
            "fields": "nextPageToken, files(id, name, mimeType, size, modifiedTime)",
            "pageSize": 100,
            "includeItemsFromAllDrives": "true",
            "supportsAllDrives": "true",
        }
        if page_token:
            params["pageToken"] = page_token

        resp = requests.get(DRIVE_FILES_URL, headers=headers, params=params, timeout=30)

        if resp.status_code != 200:
            msg = resp.json().get("error", {}).get("message", resp.text)
            raise RuntimeError(f"API error {resp.status_code}: {msg}")

        data = resp.json()

        for item in data.get("files", []):
            if stop.is_set():
                return
            if item.get("mimeType") == FOLDER_MIME:
                child_path = f"{path}{item['name']}/"
                q.put(("status", f"Scanning: {child_path}"))
                _recurse(token, item["id"], q, stop, child_path)
            else:
                q.put(("file", DriveFile(
                    folder=path,
                    name=item.get("name", ""),
                    mime=item.get("mimeType", ""),
                    size=item.get("size", "—"),
                    modified=item.get("modifiedTime", "")[:10],
                    id=item["id"],
                )))

        page_token = data.get("nextPageToken")
        if not page_token:
            break


# ── GUI ───────────────────────────────────────────────────────────────────────

C = {
    "bg":          "#1e1f22",
    "row_even":    "#2b2d30",
    "row_odd":     "#25272a",
    "head_bg":     "#1a1b1e",
    "head_fg":     "#9199a5",
    "sel_bg":      "#2d6099",
    "sel_fg":      "#ffffff",
    "fg":          "#dfe1e5",
    "fg_dim":      "#7a8189",
    "entry_bg":    "#2b2d30",
    "entry_border":"#43454a",
    "sep":         "#3c3f41",
    "btn_start":   "#2d6099",
    "btn_stop":    "#c0392b",
    "btn_export":  "#1e7e34",
    "btn_fg":      "#ffffff",
}


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Google Drive Parser")
        self.geometry("900x620")
        self.resizable(True, True)

        self._files: list[DriveFile] = []
        self._q: queue.Queue = queue.Queue()
        self._stop = threading.Event()

        self._apply_style()
        self._build_toolbar()
        self._build_table()
        self._build_statusbar()
        self._bind_macos_clipboard()

    # ── Style ─────────────────────────────────────────────────────────────────

    def _apply_style(self) -> None:
        self.configure(bg=C["bg"])

        s = ttk.Style(self)
        s.theme_use("clam")  # "clam" respects tag bg colors; "aqua" (macOS default) ignores them

        s.configure("Treeview",
            background=C["row_even"],
            foreground=C["fg"],
            fieldbackground=C["row_even"],
            rowheight=26,
            borderwidth=0,
            relief="flat",
        )
        s.configure("Treeview.Heading",
            background=C["head_bg"],
            foreground=C["head_fg"],
            font=("Helvetica Neue", 11, "bold"),
            borderwidth=0,
            relief="flat",
            padding=(6, 6),
        )
        s.map("Treeview",
            background=[("selected", C["sel_bg"])],
            foreground=[("selected", C["sel_fg"])],
        )
        s.map("Treeview.Heading",
            background=[("active", "#2e3033")],
            relief=[("active", "flat")],
        )
        s.configure("Vertical.TScrollbar",   troughcolor=C["bg"], background=C["head_bg"], borderwidth=0)
        s.configure("Horizontal.TScrollbar", troughcolor=C["bg"], background=C["head_bg"], borderwidth=0)

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build_toolbar(self) -> None:
        bar = tk.Frame(self, bg=C["bg"], pady=10, padx=12)
        bar.pack(fill="x")

        lbl_kw = dict(bg=C["bg"], fg=C["fg"], font=("Helvetica Neue", 11))
        ent_kw = dict(font=("Helvetica Neue", 11), relief="flat",
                      bg=C["entry_bg"], fg=C["fg"], insertbackground=C["fg"],
                      highlightthickness=1, highlightbackground=C["entry_border"],
                      highlightcolor=C["sel_bg"])

        tk.Label(bar, text="Access Token:", **lbl_kw).grid(row=0, column=0, sticky="w")
        self._token = tk.StringVar()
        tk.Entry(bar, textvariable=self._token, width=60, show="•", **ent_kw).grid(
            row=0, column=1, padx=(6, 16), sticky="ew", ipady=4
        )

        tk.Label(bar, text="Folder ID:", **lbl_kw).grid(row=0, column=2, sticky="w")
        self._folder = tk.StringVar()
        tk.Entry(bar, textvariable=self._folder, width=36, **ent_kw).grid(
            row=0, column=3, padx=(6, 16), ipady=4
        )

        btn_kw = dict(font=("Helvetica Neue", 11, "bold"), fg=C["btn_fg"],
                      relief="flat", padx=12, pady=4, cursor="hand2", borderwidth=0)

        self._btn_start = tk.Button(bar, text="▶  Start", command=self._on_start,
                                    bg=C["btn_start"], activebackground="#3a75b5",
                                    activeforeground=C["btn_fg"], **btn_kw)
        self._btn_start.grid(row=0, column=4, padx=(0, 6))

        self._btn_stop = tk.Button(bar, text="■  Stop", command=self._on_stop, state="disabled",
                                   bg=C["btn_stop"], activebackground="#e74c3c",
                                   activeforeground=C["btn_fg"], **btn_kw)
        self._btn_stop.grid(row=0, column=5)

        bar.columnconfigure(1, weight=1)

    def _build_table(self) -> None:
        frame = tk.Frame(self, bg=C["bg"], padx=12, pady=4)
        frame.pack(fill="both", expand=True)

        col_ids = [c[0] for c in COLUMNS]
        self._tree = ttk.Treeview(frame, columns=col_ids, show="headings", selectmode="browse")
        for key, label, width in COLUMNS:
            self._tree.heading(key, text=label, command=lambda k=key: self._sort_by(k))
            self._tree.column(key, width=width, anchor="w", minwidth=60)

        vsb = ttk.Scrollbar(frame, orient="vertical",   command=self._tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side="right",  fill="y")
        hsb.pack(side="bottom", fill="x")
        self._tree.pack(fill="both", expand=True)

        self._tree.tag_configure("even", background=C["row_even"])
        self._tree.tag_configure("odd",  background=C["row_odd"])

    def _build_statusbar(self) -> None:
        sep = tk.Frame(self, bg=C["sep"], height=1)
        sep.pack(fill="x")

        bar = tk.Frame(self, bg=C["bg"], pady=8, padx=12)
        bar.pack(fill="x")

        self._status = tk.StringVar(value="Enter token and folder ID, then press Start.")
        tk.Label(bar, textvariable=self._status, bg=C["bg"],
                 fg=C["fg_dim"], font=("Helvetica Neue", 10)).pack(side="left")

        tk.Button(bar, text="Export CSV", command=self._export_csv,
                  bg=C["btn_export"], fg=C["btn_fg"], activebackground="#27ae60",
                  activeforeground=C["btn_fg"], font=("Helvetica Neue", 11, "bold"),
                  relief="flat", padx=12, pady=4, cursor="hand2", borderwidth=0,
                  ).pack(side="right")

    # ── Actions ───────────────────────────────────────────────────────────────

    def _on_start(self) -> None:
        token     = self._token.get().strip()
        folder_id = self._folder.get().strip()

        if not token or not folder_id:
            messagebox.showwarning("Missing input", "Fill in Access Token and Folder ID.")
            return

        self._files.clear()
        self._tree.delete(*self._tree.get_children())
        self._stop.clear()
        self._set_busy(True)
        self._status.set("Starting…")

        threading.Thread(
            target=scan_folder,
            args=(token, folder_id, self._q, self._stop),
            daemon=True,
        ).start()
        self._poll()

    def _on_stop(self) -> None:
        self._stop.set()

    def _export_csv(self) -> None:
        if not self._files:
            messagebox.showinfo("No data", "Run a search first.")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV file", "*.csv")],
            initialfile="drive_files.csv",
        )
        if not path:
            return

        with open(path, "w", newline="", encoding="utf-8-sig") as fh:
            writer = csv.writer(fh)
            writer.writerow([label for _, label, _ in COLUMNS])
            writer.writerows(astuple(f) for f in self._files)

        messagebox.showinfo("Saved", f"Exported to:\n{path}")

    # ── Queue polling (main thread only) ─────────────────────────────────────

    def _poll(self) -> None:
        try:
            while True:
                kind, payload = self._q.get_nowait()
                if kind == "file":
                    self._append_row(payload)
                    self._status.set(f"Found {len(self._files)} files…")
                elif kind == "status":
                    self._status.set(payload)
                elif kind == "done":
                    self._status.set(f"Done — {len(self._files)} files found.")
                    self._set_busy(False)
                    return
                elif kind == "error":
                    messagebox.showerror("Error", payload)
                    self._status.set("Error — see dialog.")
                    self._set_busy(False)
                    return
        except queue.Empty:
            pass

        if self._stop.is_set():
            self._status.set(f"Stopped — {len(self._files)} files collected.")
            self._set_busy(False)
            return

        self.after(100, self._poll)

    # ── Table helpers ─────────────────────────────────────────────────────────

    def _append_row(self, f: DriveFile) -> None:
        self._files.append(f)
        tag = "even" if len(self._files) % 2 == 0 else "odd"
        self._tree.insert("", "end", values=astuple(f), tags=(tag,))

    def _repopulate(self) -> None:
        self._tree.delete(*self._tree.get_children())
        for i, f in enumerate(self._files):
            tag = "even" if i % 2 == 0 else "odd"
            self._tree.insert("", "end", values=astuple(f), tags=(tag,))

    def _sort_by(self, key: str) -> None:
        idx = [c[0] for c in COLUMNS].index(key)
        self._files.sort(key=lambda f: astuple(f)[idx])
        self._repopulate()

    # ── macOS clipboard ───────────────────────────────────────────────────────

    def _bind_macos_clipboard(self) -> None:
        def paste(e):
            try:
                text = e.widget.selection_get(selection="CLIPBOARD")
                try:
                    e.widget.delete("sel.first", "sel.last")
                except tk.TclError:
                    pass
                e.widget.insert("insert", text)
            except tk.TclError:
                pass
            return "break"

        def copy(e):
            try:
                text = e.widget.selection_get()
                self.clipboard_clear()
                self.clipboard_append(text)
            except tk.TclError:
                pass
            return "break"

        def cut(e):
            try:
                text = e.widget.selection_get()
                self.clipboard_clear()
                self.clipboard_append(text)
                e.widget.delete("sel.first", "sel.last")
            except tk.TclError:
                pass
            return "break"

        def select_all(e):
            e.widget.select_range(0, "end")
            e.widget.icursor("end")
            return "break"

        self.bind_class("Entry", "<Command-v>", paste)
        self.bind_class("Entry", "<Command-c>", copy)
        self.bind_class("Entry", "<Command-x>", cut)
        self.bind_class("Entry", "<Command-a>", select_all)

    # ── State ─────────────────────────────────────────────────────────────────

    def _set_busy(self, busy: bool) -> None:
        self._btn_start.config(state="disabled" if busy else "normal")
        self._btn_stop.config(state="normal"    if busy else "disabled")


if __name__ == "__main__":
    App().mainloop()
