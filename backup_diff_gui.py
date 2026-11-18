import difflib
import hashlib
import os
import queue
import subprocess
import threading
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


CHUNK_SIZE = 1024 * 1024


def format_size(num):
    if num is None:
        return ""
    try:
        num = int(num)
    except Exception:
        return str(num)
    n = float(num)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if n < 1024 or unit == "TB":
            if unit == "B":
                return f"{int(n)} {unit}"
            return f"{n:.1f} {unit}"
        n /= 1024
    return str(num)


def file_hash(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            chunk = f.read(CHUNK_SIZE)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def is_probably_text(path, sample_size=8192):
    try:
        with open(path, "rb") as f:
            sample = f.read(sample_size)
    except OSError:
        return False
    if b"\x00" in sample:
        return False
    try:
        sample.decode("utf-8")
        return True
    except UnicodeDecodeError:
        return False


def norm_rel(rel):
    if rel is None:
        return ""
    rel = str(rel).replace("/", os.sep).replace("\\", os.sep)
    rel = os.path.normpath(rel)
    if rel in (".", ""):
        return ""
    return rel.strip(os.sep)


@dataclass
class LocalSource:
    kind: str
    path: Path

    def label(self):
        return str(self.path)

    def describe(self):
        return f"Local: {self.path}"


@dataclass
class MtpSource:
    kind: str
    device: str
    storage: str
    subpath: str

    def label(self):
        sp = self.subpath.strip().strip("\\/")
        return f"{self.device}\\{self.storage}" + (f"\\{sp}" if sp else "")

    def describe(self):
        sp = self.subpath.strip().strip("\\/") or "(root)"
        return f"Phone: {self.device} / {self.storage} / {sp}"


class MtpError(Exception):
    pass


class MtpManager:
    def __init__(self, helper_path):
        self.helper_path = str(helper_path)

    def _run(self, args, timeout=600):
        cmd = [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            self.helper_path,
        ] + args
        try:
            cp = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=timeout,
            )
        except FileNotFoundError:
            raise MtpError("PowerShell was not found. Phone (MTP) support requires PowerShell on Windows.")
        except subprocess.TimeoutExpired:
            raise MtpError("Phone operation timed out.")
        if cp.returncode != 0:
            msg = (cp.stderr or cp.stdout or "Unknown PowerShell error").strip()
            raise MtpError(msg)
        # Some hosts may write warnings to stderr; ignore if success
        return cp.stdout

    def list_devices(self):
        out = self._run(["-Action", "list-devices"], timeout=60)
        items = []
        for line in out.splitlines():
            name = line.strip()
            if name:
                items.append(name)
        # de-dup preserve order
        seen = set()
        res = []
        for x in items:
            if x not in seen:
                seen.add(x)
                res.append(x)
        return res

    def list_storages(self, device):
        out = self._run(["-Action", "list-storages", "-Device", device], timeout=60)
        return [l.strip() for l in out.splitlines() if l.strip()]

    def list_children(self, device, storage, subpath=""):
        out = self._run(
            [
                "-Action",
                "list-children",
                "-Device",
                device,
                "-Storage",
                storage,
                "-SubPath",
                subpath or ".",
            ],
            timeout=120,
        )
        rows = []
        for line in out.splitlines():
            if not line.strip():
                continue
            parts = line.split("\t")
            if len(parts) < 2:
                continue
            typ = parts[0].strip()
            name = parts[1]
            size = 0
            if len(parts) >= 3:
                try:
                    size = int(parts[2])
                except Exception:
                    size = 0
            rows.append({"type": typ, "name": name, "size": size})
        return rows

    def index(self, device, storage, subpath="", progress_cb=None):
        out = self._run(
            [
                "-Action",
                "index",
                "-Device",
                device,
                "-Storage",
                storage,
                "-SubPath",
                subpath or ".",
            ],
            timeout=3600,
        )
        files = {}
        dirs = set()
        for line in out.splitlines():
            if not line.strip():
                continue
            parts = line.split("\t")
            if not parts:
                continue
            typ = parts[0].strip()
            if typ == "D" and len(parts) >= 2:
                rel = norm_rel(parts[1])
                if rel:
                    dirs.add(rel)
            elif typ == "F" and len(parts) >= 2:
                rel = norm_rel(parts[1])
                size = 0
                if len(parts) >= 3:
                    try:
                        size = int(parts[2])
                    except Exception:
                        size = 0
                if rel:
                    files[rel] = size
                    # ensure parent dirs known
                    parent = os.path.dirname(rel)
                    while parent and parent not in (".", os.sep):
                        dirs.add(parent)
                        parent = os.path.dirname(parent)
        return files, dirs

    def copy_file(self, device, storage, subpath, dest_path):
        dest_path = str(dest_path)
        Path(dest_path).parent.mkdir(parents=True, exist_ok=True)
        out = self._run(
            [
                "-Action",
                "copy-file",
                "-Device",
                device,
                "-Storage",
                storage,
                "-SubPath",
                subpath,
                "-Dest",
                dest_path,
            ],
            timeout=1800,
        )
        for line in reversed(out.splitlines()):
            p = line.strip().strip('"')
            if p:
                return p
        if os.path.exists(dest_path):
            return dest_path
        raise MtpError(f"Copy failed for {subpath}")


class SourcePanel(ttk.LabelFrame):
    def __init__(self, master, title, mtp_manager):
        super().__init__(master, text=title, padding=8)
        self.mtp = mtp_manager
        self.mode_var = tk.StringVar(value="local")

        mode_row = ttk.Frame(self)
        mode_row.grid(row=0, column=0, columnspan=4, sticky="w")
        ttk.Label(mode_row, text="Source:").pack(side="left")
        ttk.Radiobutton(mode_row, text="Local folder", variable=self.mode_var, value="local", command=self._on_mode).pack(side="left", padx=(8, 0))
        ttk.Radiobutton(mode_row, text="Phone (MTP)", variable=self.mode_var, value="mtp", command=self._on_mode).pack(side="left", padx=(8, 0))

        # Local widgets
        self.local_frame = ttk.Frame(self)
        self.local_frame.grid(row=1, column=0, columnspan=4, sticky="ew", pady=(6, 0))
        self.local_path_var = tk.StringVar()
        ttk.Label(self.local_frame, text="Folder:").grid(row=0, column=0, sticky="w")
        self.local_entry = ttk.Entry(self.local_frame, textvariable=self.local_path_var)
        self.local_entry.grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(self.local_frame, text="Browse", command=self.browse_local).grid(row=0, column=2)
        self.local_frame.columnconfigure(1, weight=1)

        # MTP widgets
        self.mtp_frame = ttk.Frame(self)
        self.mtp_frame.grid(row=2, column=0, columnspan=4, sticky="ew", pady=(6, 0))
        ttk.Label(self.mtp_frame, text="Device:").grid(row=0, column=0, sticky="w")
        self.device_var = tk.StringVar()
        self.device_combo = ttk.Combobox(self.mtp_frame, textvariable=self.device_var, state="readonly", width=30)
        self.device_combo.grid(row=0, column=1, sticky="w", padx=5)
        self.device_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_storages())

        ttk.Label(self.mtp_frame, text="Storage:").grid(row=0, column=2, sticky="w", padx=(10, 0))
        self.storage_var = tk.StringVar()
        self.storage_combo = ttk.Combobox(self.mtp_frame, textvariable=self.storage_var, state="readonly", width=24)
        self.storage_combo.grid(row=0, column=3, sticky="w", padx=5)

        ttk.Label(self.mtp_frame, text="Subfolder:").grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.subpath_var = tk.StringVar()
        ttk.Entry(self.mtp_frame, textvariable=self.subpath_var).grid(row=1, column=1, columnspan=2, sticky="ew", padx=5, pady=(4, 0))
        self.refresh_btn = ttk.Button(self.mtp_frame, text="Refresh devices", command=self.refresh_devices)
        self.refresh_btn.grid(row=1, column=3, sticky="w", padx=5, pady=(4, 0))

        self.columnconfigure(0, weight=1)
        self._on_mode()

    def _on_mode(self):
        if self.mode_var.get() == "local":
            self.mtp_frame.grid_remove()
            self.local_frame.grid()
        else:
            self.local_frame.grid_remove()
            self.mtp_frame.grid()

    def browse_local(self):
        path = filedialog.askdirectory(title="Select folder")
        if path:
            self.local_path_var.set(path)

    def refresh_devices(self):
        try:
            devices = self.mtp.list_devices()
        except Exception as e:
            messagebox.showerror("Phone (MTP)", f"Could not list devices:\n{e}")
            return
        self.device_combo["values"] = devices
        if devices and self.device_var.get() not in devices:
            self.device_var.set(devices[0])
        elif not devices:
            self.device_var.set("")
            self.storage_combo["values"] = []
            self.storage_var.set("")
        self.refresh_storages()

    def refresh_storages(self):
        dev = self.device_var.get().strip()
        if not dev:
            self.storage_combo["values"] = []
            self.storage_var.set("")
            return
        try:
            storages = self.mtp.list_storages(dev)
        except Exception as e:
            messagebox.showerror("Phone (MTP)", f"Could not list storages:\n{e}")
            storages = []
        self.storage_combo["values"] = storages
        if storages and self.storage_var.get() not in storages:
            self.storage_var.set(storages[0])
        elif not storages:
            self.storage_var.set("")

    def get_source(self):
        if self.mode_var.get() == "local":
            p = self.local_path_var.get().strip()
            return LocalSource(kind="local", path=Path(p))
        return MtpSource(
            kind="mtp",
            device=self.device_var.get().strip(),
            storage=self.storage_var.get().strip(),
            subpath=self.subpath_var.get().strip(),
        )

    def set_local_path(self, path):
        self.mode_var.set("local")
        self.local_path_var.set(str(path))
        self._on_mode()


class DiffApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Backup / Phone Folder Diff")
        self.geometry("1200x760")
        self.minsize(980, 560)

        helper = Path(__file__).with_name("mtp_helper.ps1")
        self.mtp = MtpManager(helper)

        self.filter_var = tk.StringVar(value="All")
        self.kind_filter_var = tk.StringVar(value="All")
        self.status_var = tk.StringVar(value="Select sources and click Compare.")
        self.summary_var = tk.StringVar(value="")
        self.search_var = tk.StringVar()

        self.results = []
        self.item_to_result = {}
        self.q = queue.Queue()
        self.worker = None

        self._build_ui()
        self._auto_detect_folders()
        self.search_var.trace_add("write", lambda *_: self.refresh_tree())
        self.filter_var.trace_add("write", lambda *_: self.refresh_tree())
        self.kind_filter_var.trace_add("write", lambda *_: self.refresh_tree())

    def _build_ui(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        self.left_panel = SourcePanel(top, "Side A", self.mtp)
        self.left_panel.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.right_panel = SourcePanel(top, "Side B", self.mtp)
        self.right_panel.grid(row=0, column=1, sticky="ew", padx=(6, 0))
        top.columnconfigure(0, weight=1)
        top.columnconfigure(1, weight=1)

        controls = ttk.Frame(self, padding=(10, 0, 10, 10))
        controls.pack(fill="x")

        self.compare_btn = ttk.Button(controls, text="Compare", command=self.start_compare)
        self.compare_btn.pack(side="left")
        ttk.Button(controls, text="View Diff / Details", command=self.view_selected_diff).pack(side="left", padx=(6, 0))

        ttk.Label(controls, text="Status:").pack(side="left", padx=(18, 4))
        ttk.Combobox(
            controls,
            textvariable=self.filter_var,
            values=["All", "Same", "Different", "Only in A", "Only in B"],
            state="readonly",
            width=12,
        ).pack(side="left")

        ttk.Label(controls, text="Kind:").pack(side="left", padx=(12, 4))
        ttk.Combobox(
            controls,
            textvariable=self.kind_filter_var,
            values=["All", "Files", "Folders"],
            state="readonly",
            width=10,
        ).pack(side="left")

        ttk.Label(controls, text="Search:").pack(side="left", padx=(12, 4))
        ttk.Entry(controls, textvariable=self.search_var, width=28).pack(side="left")

        self.progress = ttk.Progressbar(controls, mode="indeterminate", length=160)
        self.progress.pack(side="right")
        ttk.Label(controls, textvariable=self.status_var).pack(side="right", padx=(0, 10))

        summary_frame = ttk.Frame(self, padding=(10, 0, 10, 6))
        summary_frame.pack(fill="x")
        ttk.Label(summary_frame, textvariable=self.summary_var).pack(anchor="w")

        tree_frame = ttk.Frame(self, padding=(10, 0, 10, 10))
        tree_frame.pack(fill="both", expand=True)

        columns = ("kind", "status", "path", "size_a", "size_b")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("kind", text="Kind")
        self.tree.heading("status", text="Status")
        self.tree.heading("path", text="Relative Path")
        self.tree.heading("size_a", text="Size A")
        self.tree.heading("size_b", text="Size B")
        self.tree.column("kind", width=70, anchor="w")
        self.tree.column("status", width=110, anchor="w")
        self.tree.column("path", width=640, anchor="w")
        self.tree.column("size_a", width=110, anchor="e")
        self.tree.column("size_b", width=110, anchor="e")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", lambda e: self.view_selected_diff())

        self.tree.tag_configure("same", foreground="#2e7d32")
        self.tree.tag_configure("different", foreground="#c62828")
        self.tree.tag_configure("only", foreground="#ef6c00")

    def _auto_detect_folders(self):
        try:
            search_dirs = [Path.cwd(), Path.cwd().parent]
            candidates = []
            seen = set()
            for d in search_dirs:
                try:
                    for p in d.iterdir():
                        if p.is_dir() and "backup" in p.name.lower():
                            rp = str(p.resolve())
                            if rp not in seen:
                                seen.add(rp)
                                candidates.append(p)
                except Exception:
                    continue
            candidates.sort(key=lambda p: p.name)
            if len(candidates) >= 2:
                self.left_panel.set_local_path(candidates[0])
                self.right_panel.set_local_path(candidates[1])
            elif len(candidates) == 1:
                self.left_panel.set_local_path(candidates[0])
        except Exception:
            pass

    # ---------- source indexing ----------
    def validate_source(self, src):
        if isinstance(src, LocalSource):
            if not src.path or not Path(src.path).is_dir():
                raise ValueError(f"Local folder not found: {src.path}")
        elif isinstance(src, MtpSource):
            if not src.device:
                raise ValueError("Please select a phone device.")
            if not src.storage:
                raise ValueError("Please select a phone storage.")
        else:
            raise ValueError("Unknown source type")

    def index_local(self, root, side_label, status_cb):
        files = {}
        dirs = set()
        root = Path(root)
        for dirpath, dirnames, filenames in os.walk(root):
            dirnames.sort()
            filenames.sort()
            rel_dir = norm_rel(os.path.relpath(dirpath, root))
            if rel_dir:
                dirs.add(rel_dir)
            for name in filenames:
                full = Path(dirpath) / name
                try:
                    rel = norm_rel(full.relative_to(root))
                    st = full.stat()
                    files[rel] = {"size": st.st_size, "path": full}
                    parent = os.path.dirname(rel)
                    while parent and parent not in (".", "", os.sep):
                        dirs.add(parent)
                        parent = os.path.dirname(parent)
                except Exception:
                    continue
        return files, dirs

    def index_mtp(self, src, side_label, status_cb):
        status_cb(f"Indexing phone ({side_label})... this may take a while")
        files_raw, dirs = self.mtp.index(src.device, src.storage, src.subpath)
        files = {}
        sp = norm_rel(src.subpath)
        for rel, size in files_raw.items():
            mtp_sub = (sp + os.sep + rel) if sp else rel
            files[rel] = {
                "size": size,
                "mtp_device": src.device,
                "mtp_storage": src.storage,
                "mtp_subpath": norm_rel(mtp_sub),
            }
        return files, dirs

    # ---------- compare ----------
    def start_compare(self):
        if self.worker and self.worker.is_alive():
            return
        try:
            left_src = self.left_panel.get_source()
            right_src = self.right_panel.get_source()
            self.validate_source(left_src)
            self.validate_source(right_src)
        except Exception as e:
            messagebox.showerror("Invalid source", str(e))
            return

        self.results = []
        self.item_to_result.clear()
        self.refresh_tree()
        self.summary_var.set("")
        self.status_var.set("Preparing...")
        self.compare_btn.config(state="disabled")
        self.progress.start(10)

        self.worker = threading.Thread(
            target=self.compare_worker,
            args=(left_src, right_src),
            daemon=True,
        )
        self.worker.start()
        self.after(100, self.poll_queue)

    def files_content_equal(self, a, b):
        # Returns True/False if both files are locally accessible, else None.
        pa = a.get("path") if a else None
        pb = b.get("path") if b else None
        if pa and pb:
            try:
                return file_hash(pa) == file_hash(pb)
            except Exception:
                return False
        return None

    def files_same_for_compare(self, a, b):
        # For files with the same relative path.
        # Fast mode: no copying from phone/MTP. If content can't be checked, use size match.
        if (a.get("size") or 0) != (b.get("size") or 0):
            return False, True
        content = self.files_content_equal(a, b)
        if content is None:
            return True, False
        return bool(content), True

    def _pair_unmatched_files(self, only_a, only_b, a_files, b_files, status_cb):
        # Pair files in the same folder with the same base name but different extensions
        # (or no extension on one side). Deem them the same only if content is the same.
        # Fast mode: content is checked only when both files are locally accessible (no MTP copy).
        paired = []
        remaining_a = set(only_a)
        remaining_b = set(only_b)

        b_by_dir_stem = {}
        for rel in only_b:
            d = os.path.dirname(rel)
            base = os.path.basename(rel)
            stem, ext = os.path.splitext(base)
            b_by_dir_stem.setdefault((d, stem), []).append((rel, ext))

        for rel_a in sorted(only_a):
            if rel_a not in remaining_a:
                continue
            d = os.path.dirname(rel_a)
            base_a = os.path.basename(rel_a)
            stem, ext_a = os.path.splitext(base_a)
            cand_entries = b_by_dir_stem.get((d, stem), [])
            candidates = [rb for (rb, ext_b) in cand_entries if rb in remaining_b and ext_a != ext_b]
            if not candidates:
                continue
            a = a_files[rel_a]
            # try to find unambiguous content match among candidates
            matches = []
            for rel_b in sorted(candidates):
                b = b_files[rel_b]
                if (a.get("size") or 0) != (b.get("size") or 0):
                    continue
                status_cb(f"Checking name/extension match: {rel_a} <-> {rel_b}")
                if self.files_content_equal(a, b) is True:
                    matches.append(rel_b)
            if len(matches) == 1:
                rel_b = matches[0]
                paired.append((rel_a, rel_b))
                remaining_a.discard(rel_a)
                remaining_b.discard(rel_b)

        rem_a = sorted(remaining_a)
        rem_b = sorted(remaining_b)
        return paired, rem_a, rem_b

    def compare_worker(self, left_src, right_src):
        try:
            def status_cb(msg):
                self.q.put(("status", msg))

            # index A
            if isinstance(left_src, LocalSource):
                status_cb("Indexing A (local)...")
                a_files, a_dirs = self.index_local(left_src.path, "A", status_cb)
            else:
                a_files, a_dirs = self.index_mtp(left_src, "A", status_cb)

            # index B
            if isinstance(right_src, LocalSource):
                status_cb("Indexing B (local)...")
                b_files, b_dirs = self.index_local(right_src.path, "B", status_cb)
            else:
                b_files, b_dirs = self.index_mtp(right_src, "B", status_cb)

            in_both = sorted(set(a_files) & set(b_files))
            only_a = sorted(set(a_files) - set(b_files))
            only_b = sorted(set(b_files) - set(a_files))

            file_results = {}
            total_exact = len(in_both)
            for i, rel in enumerate(in_both, 1):
                if i == 1 or i % 20 == 0 or i == total_exact:
                    status_cb(f"Comparing files {i}/{total_exact}...")
                a = a_files[rel]
                b = b_files[rel]
                same, content_verified = self.files_same_for_compare(a, b)
                file_results[("pair", rel, rel)] = {
                    "kind": "file",
                    "status": "Same" if same else "Different",
                    "rel": rel,
                    "rel_a": rel,
                    "rel_b": rel,
                    "a": a,
                    "b": b,
                    "content_verified": content_verified,
                }

            # Pair unmatched files with same base name but different extensions (or no extension on one side)
            if only_a and only_b:
                status_cb("Matching files with same name but different extensions...")
            paired, rem_a, rem_b = self._pair_unmatched_files(only_a, only_b, a_files, b_files, status_cb)
            for rel_a, rel_b in paired:
                a = a_files[rel_a]
                b = b_files[rel_b]
                file_results[("pair", rel_a, rel_b)] = {
                    "kind": "file",
                    "status": "Same",
                    "rel": rel_a if rel_a == rel_b else f"{rel_a}  <->  {rel_b}",
                    "rel_a": rel_a,
                    "rel_b": rel_b,
                    "a": a,
                    "b": b,
                    "name_mismatch": True,
                    "content_verified": True,
                }
            for rel in rem_a:
                a = a_files[rel]
                file_results[("a", rel)] = {
                    "kind": "file",
                    "status": "Only in A",
                    "rel": rel,
                    "rel_a": rel,
                    "rel_b": None,
                    "a": a,
                    "b": None,
                }
            for rel in rem_b:
                b = b_files[rel]
                file_results[("b", rel)] = {
                    "kind": "file",
                    "status": "Only in B",
                    "rel": rel,
                    "rel_a": None,
                    "rel_b": rel,
                    "a": None,
                    "b": b,
                }

            # Determine dir statuses
            all_dirs = sorted(set(a_dirs) | set(b_dirs))
            dir_results = {}
            # collect side-specific paths of non-same file items for folder diff detection
            not_same_file_paths = {"A": set(), "B": set()}
            for r in file_results.values():
                if r["status"] != "Same":
                    if r.get("rel_a"):
                        not_same_file_paths["A"].add(r["rel_a"])
                    if r.get("rel_b"):
                        not_same_file_paths["B"].add(r["rel_b"])

            def dir_has_issue(rel):
                prefix = rel + os.sep
                for paths in (not_same_file_paths["A"], not_same_file_paths["B"]):
                    for p in paths:
                        if p == rel or p.startswith(prefix):
                            return True
                # dirs only on one side under this rel
                for d in all_dirs:
                    if d == rel or d.startswith(prefix):
                        in_a = d in a_dirs
                        in_b = d in b_dirs
                        if in_a != in_b:
                            return True
                return False

            for rel in all_dirs:
                in_a = rel in a_dirs
                in_b = rel in b_dirs
                if in_a and not in_b:
                    dir_results[rel] = {"kind": "dir", "status": "Only in A", "rel": rel, "rel_a": rel, "rel_b": None, "a": {"size": None}, "b": None}
                elif in_b and not in_a:
                    dir_results[rel] = {"kind": "dir", "status": "Only in B", "rel": rel, "rel_a": None, "rel_b": rel, "a": None, "b": {"size": None}}
                else:
                    different = dir_has_issue(rel)
                    status = "Different" if different else "Same"
                    dir_results[rel] = {"kind": "dir", "status": status, "rel": rel, "rel_a": rel, "rel_b": rel, "a": {"size": None}, "b": {"size": None}}

            results = list(dir_results.values()) + list(file_results.values())
            results.sort(key=lambda r: (0 if r["kind"] == "dir" else 1, (r.get("rel") or "").lower()))

            counts = {"Same": 0, "Different": 0, "Only in A": 0, "Only in B": 0}
            fcounts = {"Same": 0, "Different": 0, "Only in A": 0, "Only in B": 0}
            dcounts = {"Same": 0, "Different": 0, "Only in A": 0, "Only in B": 0}
            for r in results:
                counts[r["status"]] = counts.get(r["status"], 0) + 1
                if r["kind"] == "file":
                    fcounts[r["status"]] = fcounts.get(r["status"], 0) + 1
                else:
                    dcounts[r["status"]] = dcounts.get(r["status"], 0) + 1

            meta = {"counts": counts, "fcounts": fcounts, "dcounts": dcounts, "left": left_src, "right": right_src}
            self.q.put(("done", results, meta, None))
        except Exception as e:
            import traceback
            self.q.put(("done", [], {}, f"{e}\n{traceback.format_exc()}"))

    def poll_queue(self):
        try:
            while True:
                msg = self.q.get_nowait()
                kind = msg[0]
                if kind == "status":
                    self.status_var.set(msg[1])
                elif kind == "done":
                    _, results, meta, error = msg
                    self.progress.stop()
                    self.compare_btn.config(state="normal")
                    if error:
                        self.status_var.set("Error during comparison.")
                        messagebox.showerror("Comparison error", error)
                        return
                    self.results = results
                    self.refresh_tree()
                    counts = meta.get("counts", {})
                    fcounts = meta.get("fcounts", {})
                    dcounts = meta.get("dcounts", {})
                    total = sum(counts.values()) if counts else 0
                    self.summary_var.set(
                        f"Total: {total} | "
                        f"Files - Same: {fcounts.get('Same',0)}, Different: {fcounts.get('Different',0)}, "
                        f"Only in A: {fcounts.get('Only in A',0)}, Only in B: {fcounts.get('Only in B',0)} | "
                        f"Folders - Same: {dcounts.get('Same',0)}, Different: {dcounts.get('Different',0)}, "
                        f"Only in A: {dcounts.get('Only in A',0)}, Only in B: {dcounts.get('Only in B',0)}"
                    )
                    self.status_var.set("Comparison complete.")
                    return
        except queue.Empty:
            pass

        if self.worker and self.worker.is_alive():
            self.after(120, self.poll_queue)
        else:
            self.progress.stop()
            self.compare_btn.config(state="normal")

    def refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        self.item_to_result.clear()
        show = self.filter_var.get()
        kind_show = self.kind_filter_var.get()
        needle = self.search_var.get().strip().lower()

        for res in self.results:
            if show != "All" and res["status"] != show:
                continue
            if kind_show == "Files" and res["kind"] != "file":
                continue
            if kind_show == "Folders" and res["kind"] != "dir":
                continue
            if needle and needle not in res["rel"].lower():
                continue
            a = res.get("a")
            b = res.get("b")
            tag = "same" if res["status"] == "Same" else ("different" if res["status"] == "Different" else "only")
            kind_label = "Folder" if res["kind"] == "dir" else "File"
            size_a = "" if res["kind"] == "dir" else (format_size(a["size"]) if a else "")
            size_b = "" if res["kind"] == "dir" else (format_size(b["size"]) if b else "")
            item = self.tree.insert(
                "",
                "end",
                values=(kind_label, res["status"], res["rel"], size_a, size_b),
                tags=(tag,),
            )
            self.item_to_result[item] = res

    def get_selected_result(self):
        sel = self.tree.selection()
        if not sel:
            return None
        return self.item_to_result.get(sel[0])

    def view_selected_diff(self):
        res = self.get_selected_result()
        if not res:
            messagebox.showinfo("No selection", "Select an item in the list first.")
            return

        win = tk.Toplevel(self)
        win.title(f"Details - {res.get('rel', '')}")
        win.geometry("1000x720")

        header = ttk.Frame(win, padding=8)
        header.pack(fill="x")
        ttk.Label(header, text=f"Kind: {'Folder' if res['kind']=='dir' else 'File'}   Status: {res['status']}").pack(anchor="w")
        a = res.get("a")
        b = res.get("b")
        if res.get("name_mismatch") and res.get("rel_a") != res.get("rel_b"):
            ttk.Label(header, text=f"Path A: {res.get('rel_a')}").pack(anchor="w")
            ttk.Label(header, text=f"Path B: {res.get('rel_b')}").pack(anchor="w")
        ttk.Label(header, text=f"A: {self.describe_entry('A', res, a)}").pack(anchor="w")
        ttk.Label(header, text=f"B: {self.describe_entry('B', res, b)}").pack(anchor="w")

        text_frame = ttk.Frame(win, padding=(8, 0, 8, 8))
        text_frame.pack(fill="both", expand=True)
        txt = tk.Text(text_frame, wrap="none", font=("Consolas", 10))
        ysb = ttk.Scrollbar(text_frame, orient="vertical", command=txt.yview)
        xsb = ttk.Scrollbar(text_frame, orient="horizontal", command=txt.xview)
        txt.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        txt.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)

        txt.insert("1.0", "Loading...\n")
        txt.config(state="disabled")

        def worker():
            try:
                content = self.build_details_text(res)
            except Exception as e:
                import traceback
                content = f"Error building details:\n{e}\n{traceback.format_exc()}"

            def apply():
                try:
                    txt.config(state="normal")
                    txt.delete("1.0", "end")
                    txt.insert("1.0", content)
                    txt.config(state="disabled")
                except Exception:
                    pass

            self.after(0, apply)

        threading.Thread(target=worker, daemon=True).start()

    def describe_entry(self, side, res, entry):
        if not entry:
            return "(missing)"
        if res["kind"] == "dir":
            return "directory exists"
        size = format_size(entry.get("size"))
        if "path" in entry and entry["path"]:
            return f"{entry['path']} ({size})"
        if "mtp_subpath" in entry:
            return f"[Phone] {entry.get('mtp_device','')} / {entry.get('mtp_storage','')} / {entry.get('mtp_subpath','')} ({size})"
        return f"({size})"

    def _result_in_folder(self, folder_rel, r):
        prefix = folder_rel + os.sep
        for key in ("rel", "rel_a", "rel_b"):
            val = r.get(key)
            if not val:
                continue
            if val == folder_rel or val.startswith(prefix):
                return True
        return False

    def build_details_text(self, res):
        status = res["status"]
        kind = res["kind"]
        rel = res.get("rel", "")
        rel_a = res.get("rel_a") or rel
        rel_b = res.get("rel_b") or rel
        a = res.get("a")
        b = res.get("b")
        lines = []

        if kind == "dir":
            lines.append(f"Folder: {rel}")
            lines.append(f"Status: {status}")
            lines.append("")
            if status == "Only in A":
                lines.append("This folder exists only in Side A.")
            elif status == "Only in B":
                lines.append("This folder exists only in Side B.")
            elif status == "Same":
                lines.append("This folder exists on both sides and all compared contents under it match (per comparison rules).")
            else:
                lines.append("This folder exists on both sides, but some contents under it differ or are missing on one side.")
            # show child summary
            child_counts = {"Same": 0, "Different": 0, "Only in A": 0, "Only in B": 0}
            child_total = 0
            examples = []
            for r in self.results:
                if r is res:
                    continue
                if self._result_in_folder(rel, r):
                    child_total += 1
                    child_counts[r["status"]] = child_counts.get(r["status"], 0) + 1
                    if r["status"] != "Same" and len(examples) < 40:
                        examples.append(f" - [{r['kind']}] {r['status']}: {r.get('rel')}")
            lines.append("")
            lines.append(f"Items under this folder (including nested): {child_total}")
            lines.append(
                f"  Same: {child_counts.get('Same',0)}, Different: {child_counts.get('Different',0)}, "
                f"Only in A: {child_counts.get('Only in A',0)}, Only in B: {child_counts.get('Only in B',0)}"
            )
            if examples:
                lines.append("")
                lines.append("Differences/missing items (first examples):")
                lines.extend(examples)
            return "\n".join(lines) + "\n"

        # file
        lines.append(f"File: {rel}")
        lines.append(f"Status: {status}")
        if res.get("name_mismatch") and rel_a != rel_b:
            lines.append("Note: Same base name with different extension (or missing extension on one side); content matches.")
            lines.append(f"Path A: {rel_a}")
            lines.append(f"Path B: {rel_b}")
        lines.append(f"A: {self.describe_entry('A', res, a)}")
        lines.append(f"B: {self.describe_entry('B', res, b)}")
        lines.append("")

        if status == "Only in A":
            lines.append("File exists only in Side A.")
            return "\n".join(lines) + "\n"
        if status == "Only in B":
            lines.append("File exists only in Side B.")
            return "\n".join(lines) + "\n"

        a_path = a.get("path") if a else None
        b_path = b.get("path") if b else None
        content_verified = res.get("content_verified", bool(a_path and b_path))

        if status == "Same":
            if res.get("name_mismatch") and rel_a != rel_b:
                lines.append("Files are considered the same: same base name with different extension/name and identical content (SHA-256).")
            elif content_verified:
                lines.append("Files are identical (same size and SHA-256 content hash).")
            else:
                lines.append("Files are considered the same in fast mode (same path and size).")
                lines.append("Content was not byte-compared because one or both files are not locally accessible (e.g. phone/MTP) and copying is disabled.")
            return "\n".join(lines) + "\n"

        # Different
        lines.append("Files are different.")
        size_a = a.get("size", 0) if a else 0
        size_b = b.get("size", 0) if b else 0

        # Content diff/hashes only when both are local files (fast mode, no MTP copying)
        if not a_path or not b_path:
            if a_path or b_path:
                lines.append("One side is not a local file (e.g. phone/MTP); content diff is unavailable in fast mode (no copying).")
            else:
                lines.append("Neither side is a local file; content diff is unavailable in fast mode (no copying).")
            if (size_a == size_b) and not content_verified:
                lines.append("Sizes are equal but content was not verified because copying from phone/MTP is disabled.")
            return "\n".join(lines) + "\n"

        max_text_bytes = 2 * 1024 * 1024
        if size_a > max_text_bytes or size_b > max_text_bytes:
            lines.append("File too large for text diff preview.")
            try:
                lines.append(f"SHA256 A: {file_hash(a_path)}")
                lines.append(f"SHA256 B: {file_hash(b_path)}")
            except Exception as e:
                lines.append(f"Could not hash files: {e}")
            return "\n".join(lines) + "\n"

        if not is_probably_text(a_path) or not is_probably_text(b_path):
            lines.append("Binary or non-UTF8 file; text diff not shown.")
            try:
                lines.append(f"SHA256 A: {file_hash(a_path)}")
                lines.append(f"SHA256 B: {file_hash(b_path)}")
            except Exception as e:
                lines.append(f"Could not hash files: {e}")
            return "\n".join(lines) + "\n"

        try:
            with open(a_path, "r", encoding="utf-8", errors="replace") as f:
                a_lines = f.readlines()
            with open(b_path, "r", encoding="utf-8", errors="replace") as f:
                b_lines = f.readlines()
        except OSError as e:
            lines.append(f"Could not read files for diff: {e}")
            return "\n".join(lines) + "\n"

        diff = difflib.unified_diff(
            a_lines,
            b_lines,
            fromfile=str(a_path),
            tofile=str(b_path),
            lineterm="",
        )
        diff_text = "\n".join(diff)
        if not diff_text.strip():
            diff_text = "(No textual differences found, but files differ at byte/hash level.)"
        lines.append(diff_text)
        return "\n".join(lines) + "\n"


def main():
    try:
        from ctypes import windll  # type: ignore

        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    app = DiffApp()
    app.mainloop()


if __name__ == "__main__":
    main()
