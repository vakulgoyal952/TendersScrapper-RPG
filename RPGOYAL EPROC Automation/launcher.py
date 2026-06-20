"""
RPGOYAL eProcurement Tool — GUI Launcher
Cross-platform desktop app (Mac / Windows) using CustomTkinter.
Wraps both the Rajasthan eProc and National eTenders scrapers behind
a polished multi-screen interface.
"""

import os
import sys
import platform
import subprocess
import threading
import datetime
import tkinter as tk
from tkinter import messagebox

import customtkinter as ctk

# ---------------------------------------------------------------------------
# Resolve base path so imports work both from source and PyInstaller bundle
# ---------------------------------------------------------------------------

def _base_path():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_PATH = _base_path()
if BASE_PATH not in sys.path:
    sys.path.insert(0, BASE_PATH)

# ---------------------------------------------------------------------------
# Appearance
# ---------------------------------------------------------------------------
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

ACCENT = "#1F6AA5"
DARK_BG = "#1e1e2e"
LOG_FG = "#e0e0e0"
TEXT_COLOR = "#1a1a1a"
SUBTEXT_COLOR = "#666666"
FONT_FAMILY = "Segoe UI" if platform.system() == "Windows" else "Helvetica Neue"
MONO_FONT = ("Consolas", 11) if platform.system() == "Windows" else ("Menlo", 11)

# ---------------------------------------------------------------------------
# Callbacks bridge — adapts scraper callbacks to GUI-safe updates
# ---------------------------------------------------------------------------

class GUICallbacks:
    """Thread-safe bridge: scraper thread -> GUI main thread."""

    def __init__(self, app):
        self._app = app

    def on_log(self, message):
        self._app.after(0, self._app._append_log, message)

    def on_progress(self, current, total, label):
        self._app.after(0, self._app._set_progress, current, total, label)

    def on_captcha(self, label):
        event = threading.Event()
        result = {"action": "continue"}

        def _show():
            result["action"] = self._app._show_boq_action_dialog(label)
            event.set()

        self._app.after(0, _show)
        event.wait()
        return result["action"]

    def on_complete(self, summary):
        self._app.after(0, self._app._show_summary, summary)


# ---------------------------------------------------------------------------
# Main Application
# ---------------------------------------------------------------------------

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("RPGOYAL eProcurement Tool")
        self.geometry("920x660")
        self.minsize(780, 560)

        self._scraper_type = None   # "rajasthan" or "national"
        self._run_mode = "fetch"    # "fetch" or "boq_input" (rajasthan only)
        self._org_dict = {}
        self._output_root = None
        self._container = ctk.CTkFrame(self, fg_color="transparent")
        self._container.pack(fill="both", expand=True)

        self._show_welcome()

    # ===== SCREEN 1 — Welcome =====

    def _clear(self):
        for w in self._container.winfo_children():
            w.destroy()

    def _show_welcome(self):
        self._clear()
        self._scraper_type = None

        self._container.pack_propagate(False)
        outer = ctk.CTkFrame(self._container, fg_color="transparent")
        outer.pack(expand=True)

        ctk.CTkLabel(
            outer, text="RPGOYAL eProcurement Tool",
            font=(FONT_FAMILY, 28, "bold"),
            text_color=TEXT_COLOR,
        ).pack(pady=(0, 4))

        ctk.CTkLabel(
            outer,
            text="Select a portal to start scraping tenders",
            font=(FONT_FAMILY, 14),
            text_color=SUBTEXT_COLOR,
        ).pack(pady=(0, 30))

        btn_frame = ctk.CTkFrame(outer, fg_color="transparent")
        btn_frame.pack()

        ctk.CTkButton(
            btn_frame, text="Rajasthan State eProc",
            width=260, height=56, font=(FONT_FAMILY, 16, "bold"),
            command=lambda: self._on_portal_selected("rajasthan"),
        ).pack(side="left", padx=12)

        ctk.CTkButton(
            btn_frame, text="National eTenders.gov.in",
            width=260, height=56, font=(FONT_FAMILY, 16, "bold"),
            command=lambda: self._on_portal_selected("national"),
        ).pack(side="left", padx=12)

        ctk.CTkLabel(
            outer,
            text="Chrome must be installed on this machine.",
            font=(FONT_FAMILY, 11), text_color=SUBTEXT_COLOR,
        ).pack(pady=(24, 0))

    def _on_portal_selected(self, scraper_type):
        self._scraper_type = scraper_type
        self._run_mode = "fetch"
        if scraper_type == "rajasthan":
            self._show_rajasthan_mode()
        else:
            self._show_loading()

    def _show_rajasthan_mode(self):
        self._clear()

        outer = ctk.CTkFrame(self._container, fg_color="transparent")
        outer.pack(expand=True)

        ctk.CTkLabel(
            outer, text="Rajasthan eProc — choose action",
            font=(FONT_FAMILY, 22, "bold"),
            text_color=TEXT_COLOR,
        ).pack(pady=(0, 8))

        ctk.CTkLabel(
            outer,
            text="Fetch tenders from the portal, or download BOQs from your Excel list",
            font=(FONT_FAMILY, 14),
            text_color=SUBTEXT_COLOR,
        ).pack(pady=(0, 24))

        btn_frame = ctk.CTkFrame(outer, fg_color="transparent")
        btn_frame.pack()

        ctk.CTkButton(
            btn_frame, text="Fetch department-wise tenders",
            width=280, height=52, font=(FONT_FAMILY, 15, "bold"),
            command=lambda: self._on_rajasthan_mode("fetch"),
        ).pack(pady=6)

        ctk.CTkButton(
            btn_frame, text="Download BOQs from input file",
            width=280, height=52, font=(FONT_FAMILY, 15, "bold"),
            command=lambda: self._on_rajasthan_mode("boq_input"),
        ).pack(pady=6)

        ctk.CTkButton(
            outer, text="\u2190 Back", width=90, height=30,
            font=(FONT_FAMILY, 12),
            fg_color="transparent", border_width=1,
            command=self._show_welcome,
        ).pack(pady=(24, 0))

    def _on_rajasthan_mode(self, mode):
        self._run_mode = mode
        if mode == "fetch":
            self._show_loading()
        else:
            self._show_boq_input_config()

    def _show_boq_input_config(self):
        self._clear()

        import eproc_scrapper

        header = ctk.CTkFrame(self._container, fg_color="transparent")
        header.pack(fill="x", padx=24, pady=(16, 8))

        ctk.CTkButton(
            header, text="\u2190 Back", width=70, height=30,
            font=(FONT_FAMILY, 12),
            fg_color="transparent", border_width=1,
            command=self._show_rajasthan_mode,
        ).pack(side="left")

        ctk.CTkLabel(
            header, text="Download BOQs from Excel",
            font=(FONT_FAMILY, 20, "bold"),
            text_color=TEXT_COLOR,
        ).pack(side="left", padx=16)

        body = ctk.CTkFrame(self._container, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=24, pady=8)

        input_dir = eproc_scrapper.get_input_dir()
        ctk.CTkLabel(
            body,
            text=(
                "Place your .xlsx in the input folder.\n"
                "It should match exported *_tenders.xlsx (Title column with hyperlinks)."
            ),
            font=(FONT_FAMILY, 13),
            text_color=SUBTEXT_COLOR,
            justify="left",
        ).pack(anchor="w", pady=(0, 8))

        ctk.CTkLabel(
            body, text=f"Input folder:\n{input_dir}",
            font=(FONT_FAMILY, 12),
            text_color=TEXT_COLOR,
            justify="left",
        ).pack(anchor="w", pady=(0, 12))

        row = ctk.CTkFrame(body, fg_color="transparent")
        row.pack(fill="x", pady=(0, 8))
        ctk.CTkLabel(
            row, text="File name:", font=(FONT_FAMILY, 13),
            text_color=TEXT_COLOR,
        ).pack(side="left")
        self._input_file_var = ctk.StringVar(
            value=eproc_scrapper.DEFAULT_INPUT_XLSX
        )
        ctk.CTkEntry(
            row, textvariable=self._input_file_var, width=280, height=32,
            placeholder_text="selected_tenders.xlsx",
        ).pack(side="left", padx=(8, 0))

        ctk.CTkLabel(
            body,
            text="Output: results/rajasthan/<date>/BOQs/Miscellaneous BoQs/",
            font=(FONT_FAMILY, 12),
            text_color=SUBTEXT_COLOR,
        ).pack(anchor="w", pady=(8, 16))

        ctk.CTkButton(
            body, text="Start BOQ download", height=44,
            font=(FONT_FAMILY, 15, "bold"),
            command=self._start_boq_from_input,
        ).pack(pady=(8, 4))

    def _start_boq_from_input(self):
        import eproc_scrapper

        filename = self._input_file_var.get().strip()
        path = eproc_scrapper.resolve_input_xlsx_path(filename)
        if not os.path.isfile(path):
            messagebox.showerror(
                "File not found",
                f"Could not find:\n{path}\n\nCopy your Excel file into the input folder.",
            )
            return
        config = {
            "mode": "boq_input",
            "input_xlsx_path": path,
        }
        self._show_progress(config, title="BOQ download — input file")

    # ===== SCREEN 2 — Loading organisations =====

    def _show_loading(self):
        self._clear()

        frame = ctk.CTkFrame(self._container, fg_color="transparent")
        frame.pack(expand=True)

        portal_name = (
            "eproc.rajasthan.gov.in"
            if self._scraper_type == "rajasthan"
            else "etenders.gov.in"
        )
        ctk.CTkLabel(
            frame, text=f"Connecting to {portal_name}...",
            font=(FONT_FAMILY, 18),
            text_color=TEXT_COLOR,
        ).pack(pady=(0, 16))

        self._load_bar = ctk.CTkProgressBar(frame, width=340, mode="indeterminate")
        self._load_bar.pack()
        self._load_bar.start()

        ctk.CTkLabel(
            frame,
            text="Fetching organisation list (headless Chrome)",
            font=(FONT_FAMILY, 12), text_color=SUBTEXT_COLOR,
        ).pack(pady=(12, 0))

        threading.Thread(target=self._fetch_orgs_worker, daemon=True).start()

    def _fetch_orgs_worker(self):
        try:
            if self._scraper_type == "rajasthan":
                import eproc_scrapper
                org_dict = eproc_scrapper.fetch_organisations()
            else:
                import national_eproc_scrapper
                org_dict = national_eproc_scrapper.fetch_organisations()
            self.after(0, self._on_orgs_loaded, org_dict, None)
        except Exception as e:
            self.after(0, self._on_orgs_loaded, {}, str(e))

    def _on_orgs_loaded(self, org_dict, error):
        self._load_bar.stop()
        if error or not org_dict:
            msg = error or "No organisations found on the portal."
            messagebox.showerror("Connection Error", msg)
            self._show_welcome()
            return
        self._org_dict = org_dict
        self._show_config()

    # ===== SCREEN 3 — Configuration form =====

    def _show_config(self):
        self._clear()

        title_text = (
            "Rajasthan eProc" if self._scraper_type == "rajasthan"
            else "National eTenders"
        )
        header = ctk.CTkFrame(self._container, fg_color="transparent")
        header.pack(fill="x", padx=24, pady=(16, 8))

        ctk.CTkButton(
            header, text="\u2190 Back", width=70, height=30,
            font=(FONT_FAMILY, 12),
            fg_color="transparent", border_width=1,
            command=self._show_welcome,
        ).pack(side="left")

        ctk.CTkLabel(
            header, text=title_text,
            font=(FONT_FAMILY, 20, "bold"),
            text_color=TEXT_COLOR,
        ).pack(side="left", padx=16)

        body = ctk.CTkFrame(self._container, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=24, pady=8)

        # --- Organisation selector ---
        ctk.CTkLabel(
            body, text="Select Organisations",
            font=(FONT_FAMILY, 14, "bold"),
            text_color=TEXT_COLOR,
        ).pack(anchor="w")

        search_frame = ctk.CTkFrame(body, fg_color="transparent")
        search_frame.pack(fill="x", pady=(4, 4))

        self._search_var = ctk.StringVar()
        self._search_var.trace_add("write", self._filter_org_list)
        ctk.CTkEntry(
            search_frame, textvariable=self._search_var,
            placeholder_text="Search organisations...",
            width=320, height=32,
        ).pack(side="left")

        ctk.CTkButton(
            search_frame, text="Select All", width=90, height=30,
            font=(FONT_FAMILY, 11),
            command=lambda: self._toggle_all_orgs(True),
        ).pack(side="left", padx=(12, 4))

        ctk.CTkButton(
            search_frame, text="Clear All", width=90, height=30,
            font=(FONT_FAMILY, 11),
            fg_color="transparent", border_width=1,
            command=lambda: self._toggle_all_orgs(False),
        ).pack(side="left", padx=4)

        list_frame = ctk.CTkScrollableFrame(body, height=200)
        list_frame.pack(fill="both", expand=True, pady=(0, 8))
        self._org_list_frame = list_frame

        self._org_vars = {}
        for name in sorted(self._org_dict.keys()):
            var = tk.BooleanVar(value=False)
            self._org_vars[name] = var

        self._org_widgets = {}
        self._render_org_checkboxes()

        # --- Options row ---
        opts = ctk.CTkFrame(body, fg_color="transparent")
        opts.pack(fill="x", pady=(4, 4))

        # Value range
        val_frame = ctk.CTkFrame(opts, fg_color="transparent")
        val_frame.pack(side="left")

        ctk.CTkLabel(val_frame, text="Value Range (Crore):", font=(FONT_FAMILY, 13), text_color=TEXT_COLOR).pack(side="left")

        self._all_values_var = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            val_frame, text="All", variable=self._all_values_var,
            font=(FONT_FAMILY, 12),
            command=self._toggle_value_fields,
        ).pack(side="left", padx=(8, 8))

        self._min_val_entry = ctk.CTkEntry(val_frame, width=70, height=30, placeholder_text="Min")
        self._min_val_entry.pack(side="left", padx=2)
        ctk.CTkLabel(val_frame, text="to", font=(FONT_FAMILY, 12), text_color=TEXT_COLOR).pack(side="left", padx=4)
        self._max_val_entry = ctk.CTkEntry(val_frame, width=70, height=30, placeholder_text="Max")
        self._max_val_entry.pack(side="left", padx=2)

        self._min_val_entry.configure(state="disabled")
        self._max_val_entry.configure(state="disabled")

        # BOQ toggle
        self._boq_var = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            opts, text="Download BOQs", variable=self._boq_var,
            font=(FONT_FAMILY, 13),
        ).pack(side="left", padx=(24, 0))

        # Workers (national only)
        if self._scraper_type == "national":
            ctk.CTkLabel(
                opts, text="Workers:", font=(FONT_FAMILY, 13),
                text_color=TEXT_COLOR,
            ).pack(side="left", padx=(24, 4))
            self._workers_var = tk.IntVar(value=4)
            ctk.CTkSlider(
                opts, from_=1, to=8, number_of_steps=7,
                variable=self._workers_var, width=120,
            ).pack(side="left")
            self._workers_label = ctk.CTkLabel(
                opts, text="4", font=(FONT_FAMILY, 13, "bold"), width=24,
            )
            self._workers_label.pack(side="left", padx=(4, 0))
            self._workers_var.trace_add(
                "write",
                lambda *_: self._workers_label.configure(
                    text=str(self._workers_var.get())
                ),
            )

        # Start button
        ctk.CTkButton(
            body, text="Start Scraping", height=44,
            font=(FONT_FAMILY, 15, "bold"),
            command=self._start_scraping,
        ).pack(pady=(12, 4))

    def _render_org_checkboxes(self, filter_text=""):
        for w in self._org_list_frame.winfo_children():
            w.destroy()
        self._org_widgets = {}
        ft = filter_text.lower()
        for name in sorted(self._org_dict.keys()):
            if ft and ft not in name.lower():
                continue
            cb = ctk.CTkCheckBox(
                self._org_list_frame, text=name,
                variable=self._org_vars[name],
                font=(FONT_FAMILY, 12),
            )
            cb.pack(anchor="w", pady=1)
            self._org_widgets[name] = cb

    def _filter_org_list(self, *_args):
        self._render_org_checkboxes(self._search_var.get())

    def _toggle_all_orgs(self, state):
        ft = self._search_var.get().lower()
        for name, var in self._org_vars.items():
            if not ft or ft in name.lower():
                var.set(state)

    def _toggle_value_fields(self):
        state = "disabled" if self._all_values_var.get() else "normal"
        self._min_val_entry.configure(state=state)
        self._max_val_entry.configure(state=state)

    def _start_scraping(self):
        selected = [
            (name, self._org_dict[name])
            for name, var in self._org_vars.items() if var.get()
        ]
        if not selected:
            messagebox.showwarning("No Selection", "Please select at least one organisation.")
            return

        value_mode = "all"
        range_min_r = None
        range_max_r = None
        if not self._all_values_var.get():
            try:
                lo = float(self._min_val_entry.get())
                hi = float(self._max_val_entry.get())
            except ValueError:
                messagebox.showwarning("Invalid Range", "Enter valid numbers for min and max crore.")
                return
            if lo <= 0 or hi <= 0:
                messagebox.showwarning("Invalid Range", "Crore values must be positive.")
                return
            if lo > hi:
                lo, hi = hi, lo
            value_mode = "range"
            range_min_r = int(round(lo * 10_000_000))
            range_max_r = int(round(hi * 10_000_000))

        config = {
            "org_jobs": selected,
            "value_mode": value_mode,
            "range_min_r": range_min_r,
            "range_max_r": range_max_r,
            "download_boq": self._boq_var.get(),
        }
        if self._scraper_type == "national":
            config["num_workers"] = self._workers_var.get()

        self._show_progress(config)

    def _show_progress(self, config, title=None):
        self._clear()

        header = ctk.CTkFrame(self._container, fg_color="transparent")
        header.pack(fill="x", padx=24, pady=(16, 4))

        if title:
            header_text = title
        elif config.get("mode") == "boq_input":
            header_text = "BOQ download — input file"
        else:
            portal = (
                "Rajasthan eProc" if self._scraper_type == "rajasthan"
                else "National eTenders"
            )
            header_text = f"Scraping — {portal}"
        ctk.CTkLabel(
            header, text=header_text,
            font=(FONT_FAMILY, 20, "bold"),
            text_color=TEXT_COLOR,
        ).pack(side="left")

        body = ctk.CTkFrame(self._container, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=24, pady=8)

        self._status_label = ctk.CTkLabel(
            body, text="Initialising...",
            font=(FONT_FAMILY, 14),
            text_color=TEXT_COLOR,
        )
        self._status_label.pack(anchor="w", pady=(0, 4))

        self._progress_bar = ctk.CTkProgressBar(body, width=400)
        self._progress_bar.pack(fill="x", pady=(0, 4))
        self._progress_bar.set(0)

        self._pct_label = ctk.CTkLabel(
            body, text="", font=(FONT_FAMILY, 12), text_color=SUBTEXT_COLOR,
        )
        self._pct_label.pack(anchor="w", pady=(0, 8))

        self._log_box = ctk.CTkTextbox(
            body, font=MONO_FONT,
            fg_color=DARK_BG, text_color=LOG_FG,
            state="disabled",
        )
        self._log_box.pack(fill="both", expand=True)

        callbacks = GUICallbacks(self)
        threading.Thread(
            target=self._run_scraper_worker,
            args=(config, callbacks),
            daemon=True,
        ).start()

    def _run_scraper_worker(self, config, callbacks):
        try:
            if config.get("mode") == "boq_input":
                import eproc_scrapper
                output_root, summary = eproc_scrapper.run_boq_from_input_file(
                    config, callbacks
                )
            elif self._scraper_type == "rajasthan":
                import eproc_scrapper
                output_root, summary = eproc_scrapper.run_scraper(config, callbacks)
            else:
                import national_eproc_scrapper
                output_root, summary = national_eproc_scrapper.run_scraper(
                    config, callbacks
                )
            self._output_root = output_root
        except Exception as e:
            self.after(0, self._append_log, f"ERROR: {e}")
            self.after(0, self._show_summary, [])

    def _append_log(self, message):
        self._log_box.configure(state="normal")
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        self._log_box.insert("end", f"[{ts}] {message}\n")
        self._log_box.see("end")
        self._log_box.configure(state="disabled")

    def _set_progress(self, current, total, label):
        if total > 0:
            frac = current / total
            self._progress_bar.set(frac)
            self._pct_label.configure(
                text=f"{current}/{total} — {round(frac * 100, 1)}%"
            )
        self._status_label.configure(text=label)

    def _show_boq_action_dialog(self, label):
        dialog = ctk.CTkToplevel(self)
        dialog.title("BOQ Action Required")
        dialog.geometry("520x260")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        dialog.lift()
        dialog.focus_force()

        result = {"action": "continue"}

        ctk.CTkLabel(
            dialog, text="BOQ download needs your help",
            font=(FONT_FAMILY, 18, "bold"),
            text_color=TEXT_COLOR,
        ).pack(pady=(20, 8))

        ctk.CTkLabel(
            dialog, text=label,
            font=(FONT_FAMILY, 12), wraplength=460,
            text_color=TEXT_COLOR,
        ).pack(pady=(0, 4))

        ctk.CTkLabel(
            dialog,
            text=(
                "Continue: page/CAPTCHA fixed — retry in the same Chrome window.\n"
                "Restart browser: close Chrome and retry THIS tender in a new window.\n"
                "Skip: move to the next tender."
            ),
            font=(FONT_FAMILY, 12), text_color=SUBTEXT_COLOR,
            justify="left",
        ).pack(pady=(0, 16))

        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack()

        def _continue():
            result["action"] = "continue"
            dialog.destroy()

        def _restart():
            result["action"] = "restart"
            dialog.destroy()

        def _skip():
            result["action"] = "skip"
            dialog.destroy()

        ctk.CTkButton(
            btn_frame, text="Continue", width=110, height=36,
            font=(FONT_FAMILY, 13, "bold"),
            command=_continue,
        ).pack(side="left", padx=6)

        ctk.CTkButton(
            btn_frame, text="Restart browser", width=130, height=36,
            font=(FONT_FAMILY, 13, "bold"),
            fg_color="#C65D00",
            hover_color="#A04C00",
            command=_restart,
        ).pack(side="left", padx=6)

        ctk.CTkButton(
            btn_frame, text="Skip", width=90, height=36,
            font=(FONT_FAMILY, 13),
            fg_color="transparent", border_width=1,
            command=_skip,
        ).pack(side="left", padx=6)

        dialog.wait_window()
        return result["action"]

    # ===== SCREEN 5 — Summary =====

    def _show_summary(self, summary):
        self._clear()

        outer = ctk.CTkFrame(self._container, fg_color="transparent")
        outer.pack(fill="both", expand=True, padx=24, pady=16)

        ctk.CTkLabel(
            outer, text="Scraping Complete",
            font=(FONT_FAMILY, 24, "bold"),
            text_color=TEXT_COLOR,
        ).pack(pady=(0, 16))

        if summary:
            table_frame = ctk.CTkFrame(outer)
            table_frame.pack(fill="x", pady=(0, 16))

            headers = ["Organisation", "Total", "After Filter", "BOQs"]
            for col, h in enumerate(headers):
                lbl = ctk.CTkLabel(
                    table_frame, text=h,
                    font=(FONT_FAMILY, 12, "bold"),
                    text_color=TEXT_COLOR,
                    anchor="w" if col == 0 else "center",
                )
                lbl.grid(row=0, column=col, padx=8, pady=(8, 4), sticky="ew")

            for r, s in enumerate(summary, start=1):
                vals = [
                    s.get("org", ""),
                    str(s.get("total", 0)),
                    str(s.get("filtered", 0)),
                    str(s.get("boqs", 0)),
                ]
                for col, v in enumerate(vals):
                    lbl = ctk.CTkLabel(
                        table_frame, text=v,
                        font=(FONT_FAMILY, 12),
                        text_color=TEXT_COLOR,
                        anchor="w" if col == 0 else "center",
                    )
                    lbl.grid(row=r, column=col, padx=8, pady=2, sticky="ew")

            table_frame.grid_columnconfigure(0, weight=3)
            for c in range(1, 4):
                table_frame.grid_columnconfigure(c, weight=1)
        else:
            ctk.CTkLabel(
                outer, text="No tenders were exported.",
                font=(FONT_FAMILY, 14), text_color=SUBTEXT_COLOR,
            ).pack(pady=8)

        if self._output_root and os.path.isdir(self._output_root):
            path_frame = ctk.CTkFrame(outer, fg_color="transparent")
            path_frame.pack(fill="x", pady=(0, 8))

            ctk.CTkLabel(
                path_frame, text="Output folder:",
                font=(FONT_FAMILY, 13),
                text_color=TEXT_COLOR,
            ).pack(side="left")

            ctk.CTkLabel(
                path_frame, text=self._output_root,
                font=(FONT_FAMILY, 12), text_color=ACCENT,
            ).pack(side="left", padx=(8, 0))

            ctk.CTkButton(
                outer, text="Open Results Folder", width=180, height=36,
                font=(FONT_FAMILY, 13),
                command=self._open_output_folder,
            ).pack(pady=(0, 8))

        btn_row = ctk.CTkFrame(outer, fg_color="transparent")
        btn_row.pack(pady=(16, 0))

        ctk.CTkButton(
            btn_row, text="New Scrape", width=140, height=40,
            font=(FONT_FAMILY, 14, "bold"),
            command=self._show_welcome,
        ).pack(side="left", padx=8)

        ctk.CTkButton(
            btn_row, text="Quit", width=120, height=40,
            font=(FONT_FAMILY, 14),
            fg_color="transparent", border_width=1,
            command=self.destroy,
        ).pack(side="left", padx=8)

    def _open_output_folder(self):
        path = self._output_root
        if not path or not os.path.isdir(path):
            return
        system = platform.system()
        if system == "Darwin":
            subprocess.Popen(["open", path])
        elif system == "Windows":
            os.startfile(path)
        else:
            subprocess.Popen(["xdg-open", path])


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app = App()
    app.mainloop()
