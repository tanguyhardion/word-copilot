import os
import sys
import threading
import tkinter as tk
from tkinter import scrolledtext, messagebox

from word_copilot.colors import refresh_word_theme_colors
from word_copilot.connector import WordConnector
from word_copilot.dsl.edit_parser import parse_edit_dsl, looks_like_edit_dsl
from word_copilot.dsl.parser import parse_dsl_pages
from word_copilot.editor import WordEditor
from word_copilot.extractor import WordExtractor
from word_copilot.gui.highlighter import DSLHighlighter
from word_copilot.gui.sample import SAMPLE_DSL


def resource_path(relative_path):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Word Copilot")
        self.root.state("zoomed")
        self.root.configure(bg="#1E1E2E")
        try:
            icon_path = resource_path("assets/icon.ico")
            self.root.iconbitmap(icon_path)
        except Exception:
            pass
        self.connector = WordConnector()
        self.extractor = WordExtractor()
        self.editor = WordEditor()
        self._build()

    def _build(self):
        BG_MAIN = "#1E1E2E"
        BG_PANEL = "#2B2B2B"
        FG_MAIN = "#F3F3F3"
        FG_MUTED = "#B8B8B8"
        WORD_BLUE = "#2B579A"
        WORD_BLUE_HOVER = "#3A6DB5"
        WORD_TEAL = "#217346"
        WORD_TEAL_HOVER = "#2E8F5C"
        WORD_GRAY = "#4A4F57"
        WORD_GRAY_HOVER = "#5A606A"
        WORD_AMBER = "#8B5E00"
        WORD_AMBER_HOVER = "#A87000"
        WORD_PLUM = "#5C3D8B"
        WORD_PLUM_HOVER = "#71499E"

        # ── header ────────────────────────────────────────────────────────────
        header = tk.Frame(self.root, bg=BG_PANEL, height=48)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(
            header,
            text="📄  Word Copilot",
            font=("Segoe UI", 14, "bold"),
            bg=BG_PANEL,
            fg=FG_MAIN,
        ).pack(side="left", padx=16, pady=10)

        # ── editor ────────────────────────────────────────────────────────────
        editor_frame = tk.Frame(self.root, bg=BG_MAIN)
        editor_frame.pack(fill="both", expand=True, padx=10, pady=(8, 4))
        tk.Label(
            editor_frame,
            text="Paste DSL below (normal build DSL, or 'edit target=active …' edit-mode DSL):",
            font=("Segoe UI", 9),
            bg=BG_MAIN,
            fg=FG_MUTED,
        ).pack(anchor="w")

        self.text_editor = scrolledtext.ScrolledText(
            editor_frame,
            wrap="word",
            font=("Consolas", 10),
            bg=BG_MAIN,
            fg="#CDD6F4",
            insertbackground=WORD_BLUE,
            selectbackground="#3E5F99",
            selectforeground="#FFFFFF",
            relief="flat",
            bd=8,
            undo=True,
        )
        self.text_editor.pack(fill="both", expand=True)
        self.text_editor.insert("1.0", SAMPLE_DSL)

        self.highlighter = DSLHighlighter(self.text_editor)
        self.text_editor.bind("<KeyRelease>", self.highlighter.highlight)

        # ── button row ────────────────────────────────────────────────────────
        btn_row = tk.Frame(self.root, bg=BG_MAIN)
        btn_row.pack(fill="x", padx=10, pady=(4, 6))

        self.open_btn = tk.Button(
            btn_row,
            text="📂  Build & Open in Word",
            font=("Segoe UI", 11, "bold"),
            bg=WORD_BLUE,
            fg="white",
            activebackground=WORD_BLUE_HOVER,
            activeforeground="white",
            relief="flat",
            bd=0,
            padx=18,
            pady=8,
            cursor="hand2",
            command=self.build_and_open,
        )
        self.open_btn.pack(side="left", padx=(0, 8))

        self.insert_btn = tk.Button(
            btn_row,
            text="➕  Insert at Cursor",
            font=("Segoe UI", 11, "bold"),
            bg=WORD_TEAL,
            fg="white",
            activebackground=WORD_TEAL_HOVER,
            activeforeground="white",
            relief="flat",
            bd=0,
            padx=18,
            pady=8,
            cursor="hand2",
            command=self.insert_at_cursor,
        )
        self.insert_btn.pack(side="left", padx=(0, 8))

        self.save_btn = tk.Button(
            btn_row,
            text="💾  Save .docx",
            font=("Segoe UI", 11, "bold"),
            bg=WORD_GRAY,
            fg="white",
            activebackground=WORD_GRAY_HOVER,
            activeforeground="white",
            relief="flat",
            bd=0,
            padx=18,
            pady=8,
            cursor="hand2",
            command=self.save_docx,
        )
        self.save_btn.pack(side="left", padx=(0, 8))

        self.extract_btn = tk.Button(
            btn_row,
            text="🔍  Extract DSL from Word",
            font=("Segoe UI", 11, "bold"),
            bg=WORD_AMBER,
            fg="white",
            activebackground=WORD_AMBER_HOVER,
            activeforeground="white",
            relief="flat",
            bd=0,
            padx=18,
            pady=8,
            cursor="hand2",
            command=self.extract_from_word,
        )
        self.extract_btn.pack(side="left", padx=(0, 8))

        self.apply_edits_btn = tk.Button(
            btn_row,
            text="✏️  Apply Edits to Word",
            font=("Segoe UI", 11, "bold"),
            bg=WORD_PLUM,
            fg="white",
            activebackground=WORD_PLUM_HOVER,
            activeforeground="white",
            relief="flat",
            bd=0,
            padx=18,
            pady=8,
            cursor="hand2",
            command=self.apply_edits,
        )
        self.apply_edits_btn.pack(side="left", padx=(0, 8))

        # ── status bar ────────────────────────────────────────────────────────
        self.status = tk.Label(
            self.root,
            text="Ready — paste DSL and hit a button",
            font=("Segoe UI", 9),
            bg=BG_PANEL,
            fg=FG_MUTED,
            anchor="w",
        )
        self.status.pack(fill="x", side="bottom", ipady=5, padx=10)

        self.root.after(100, self.highlighter.highlight)

    def _set_status(self, msg, level="info"):
        colors = {
            "info": "#B8B8B8",
            "success": "#70AD47",
            "error": "#C55A5A",
            "warning": "#FFC000",
        }
        self.status.config(text=msg, fg=colors.get(level, "#B8B8B8"))

    def _get_pages(self):
        text = self.text_editor.get("1.0", "end").strip()
        if not text:
            return [[]]
        refresh_word_theme_colors()
        return parse_dsl_pages(text)

    def _get_edit_ops(self):
        text = self.text_editor.get("1.0", "end").strip()
        refresh_word_theme_colors()
        return parse_edit_dsl(text)

    def _lock(self):
        for btn in (
            self.open_btn,
            self.insert_btn,
            self.save_btn,
            self.extract_btn,
            self.apply_edits_btn,
        ):
            btn.config(state="disabled", text="⏳  Working…")

    def _unlock(self):
        self.open_btn.config(state="normal", text="📂  Build & Open in Word")
        self.insert_btn.config(state="normal", text="➕  Insert at Cursor")
        self.save_btn.config(state="normal", text="💾  Save .docx")
        self.extract_btn.config(state="normal", text="🔍  Extract DSL from Word")
        self.apply_edits_btn.config(state="normal", text="✏️  Apply Edits to Word")

    # ── build & open ──────────────────────────────────────────────────────────
    def build_and_open(self):
        try:
            pages = self._get_pages()
        except Exception as e:
            messagebox.showerror("Parse Error", str(e))
            return
        if not any(pages):
            messagebox.showwarning("Empty", "Nothing to build.")
            return
        self._lock()

        def run():
            try:
                self.connector.build_and_open(
                    pages,
                    status_cb=lambda m: self.root.after(0, self._set_status, m, "info"),
                )
                self.root.after(0, self._unlock)
                self.root.after(
                    0,
                    self._set_status,
                    "✓ Document built and opened in Word!",
                    "success",
                )
            except Exception as e:
                self.root.after(0, self._on_error, str(e))

        threading.Thread(target=run, daemon=True).start()

    # ── insert at cursor ──────────────────────────────────────────────────────
    def insert_at_cursor(self):
        try:
            pages = self._get_pages()
        except Exception as e:
            messagebox.showerror("Parse Error", str(e))
            return
        if not any(pages):
            messagebox.showwarning("Empty", "Nothing to insert.")
            return
        self._lock()

        def run():
            try:
                self.connector.insert_at_cursor(
                    pages,
                    status_cb=lambda m: self.root.after(0, self._set_status, m, "info"),
                )
                self.root.after(0, self._unlock)
                self.root.after(
                    0, self._set_status, "✓ Content inserted at cursor!", "success"
                )
            except Exception as e:
                self.root.after(0, self._on_error, str(e))

        threading.Thread(target=run, daemon=True).start()

    # ── save .docx ────────────────────────────────────────────────────────────
    def save_docx(self):
        try:
            pages = self._get_pages()
        except Exception as e:
            messagebox.showerror("Parse Error", str(e))
            return
        if not any(pages):
            messagebox.showwarning("Empty", "Nothing to save.")
            return
        self._lock()

        def run():
            try:
                path = self.connector.save_docx(
                    pages,
                    status_cb=lambda m: self.root.after(0, self._set_status, m, "info"),
                )
                self.root.after(0, self._unlock)
                self.root.after(0, self._set_status, f"✓ Saved to {path}", "success")
            except Exception as e:
                self.root.after(0, self._on_error, str(e))

        threading.Thread(target=run, daemon=True).start()

    # ── extract DSL from Word ─────────────────────────────────────────────────
    def extract_from_word(self):
        current = self.text_editor.get("1.0", "end").strip()
        if current:
            ok = messagebox.askyesno(
                "Replace editor content?",
                "The editor currently has DSL content.\n\n"
                "Extracting from Word will REPLACE it.\n\n"
                "Continue?",
            )
            if not ok:
                return

        self._lock()

        def run():
            try:
                self.root.after(0, self._set_status, "Connecting to Word…", "info")
                dsl = self.extractor.extract(
                    status_cb=lambda m: self.root.after(0, self._set_status, m, "info"),
                )

                def update_editor():
                    self.text_editor.delete("1.0", "end")
                    self.text_editor.insert("1.0", dsl)
                    self.highlighter.highlight()
                    self._unlock()
                    self._set_status(
                        "✓ DSL extracted from Word (with // id=N tags) — "
                        "edit elements directly, or write 'edit target=active' "
                        "ops and click Apply Edits.",
                        "success",
                    )

                self.root.after(0, update_editor)
            except Exception as e:
                self.root.after(0, self._on_error, str(e))

        threading.Thread(target=run, daemon=True).start()

    # ── apply edits (delete / replace / insert) to active Word doc ───────────
    def apply_edits(self):
        text = self.text_editor.get("1.0", "end").strip()
        if not text:
            messagebox.showwarning("Empty", "Nothing to apply.")
            return
        if not looks_like_edit_dsl(text):
            ok = messagebox.askyesno(
                "Not edit-mode DSL",
                "This doesn't look like edit-mode DSL (expected it to start "
                "with 'edit target=active' followed by delete/replace/"
                "insert_* ops).\n\n"
                "Try to parse it as edit ops anyway?",
            )
            if not ok:
                return
        try:
            ops = self._get_edit_ops()
        except Exception as e:
            messagebox.showerror("Parse Error", str(e))
            return
        if not ops:
            messagebox.showwarning("Empty", "No edit operations found.")
            return

        self._lock()

        def run():
            try:
                self.editor.apply(
                    ops,
                    status_cb=lambda m: self.root.after(0, self._set_status, m, "info"),
                )
                self.root.after(0, self._unlock)
                self.root.after(
                    0,
                    self._set_status,
                    f"✓ Applied {len(ops)} edit(s) to the document!",
                    "success",
                )
            except Exception as e:
                self.root.after(0, self._on_error, str(e))

        threading.Thread(target=run, daemon=True).start()

    def _on_error(self, msg):
        self._unlock()
        self._set_status(f"Error: {msg}", "error")
        if "win32com" in msg or "No module named 'docx'" in msg or "No module" in msg:
            messagebox.showerror(
                "Missing Dependency",
                "Required packages:\n\n"
                "  pip install pywin32 python-docx requests cairosvg\n\n"
                "Microsoft Word must also be installed.",
            )
        else:
            messagebox.showerror("Error", msg)

    def run(self):
        self.root.mainloop()
