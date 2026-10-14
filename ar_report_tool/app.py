"""AR Outstanding Report -- desktop app.

A 3-step wizard: Add reports -> Generate -> Review results.

Step 1 collects the two required source files plus an optional previous AR
report, and the output folder. Step 2 runs build_report() in a background
thread. Step 3 shows whatever came back:

- Success: the Public Ins / TPA & CPR / OP GNL totals, with a button to open
  the output folder.
- NEEDS REVIEW: a bill landed in TPA & CPR under a category this tool
  doesn't recognise -- grouped by Category + Company Name (since bills can
  share a category but belong to different payers), each group gets a
  dropdown to pick where it actually belongs. Applying regenerates the
  report; decisions are remembered in category_overrides.json so the same
  category/company pair is never flagged again.
- Missing/renamed columns: a report's columns don't match this tool's
  built-in alias list (a template change). Each missing field gets a
  dropdown of the file's actual headers to map it to. Applying retries;
  mappings are remembered in column_overrides.json.
- Failure: anything else that went wrong, shown in full so it can be
  diagnosed, with a way back to Step 1 to pick different files.

This is the file PyInstaller packages into the .exe -- see BUILD.md /
README.md. Needs `customtkinter` (see requirements.txt); when packaging,
include its data files: `pyinstaller --collect-data=customtkinter ...`.
"""
from __future__ import annotations

import json
import os
import sys
import threading
import traceback
from pathlib import Path
from tkinter import StringVar, filedialog, messagebox

import customtkinter as ctk
import pandas as pd

from logic import (
    BuildResult, FIELD_LABELS, MissingColumnsError, SourceFormatError,
    build_report, override_key,
)

APP_TITLE = "AR Outstanding Report Builder"

APP_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
REPORTS_DIR = APP_DIR / "Reports"
CATEGORY_OVERRIDES_PATH = APP_DIR / "category_overrides.json"
COLUMN_OVERRIDES_PATH = APP_DIR / "column_overrides.json"

BUCKET_CHOICES = {
    "Keep as TPA & CPR": "TPA",
    "Move to Public Ins": "PUBLIC",
    "Move to OP GNL": "OP GNL",
}

REPORT_LABEL_TO_KEY = {
    "Total Bill Wise Report": "OP",
    "Discharge Wise IP Revenue Report": "IP",
}

# -- palette, loosely matching the approved design mockup ------------------
BG = "#EEF2F8"
CARD_BG = "#FFFFFF"
BORDER = "#E3E8F0"
FIELD_BORDER = "#D4DAE3"
PRIMARY = "#1B4F91"
PRIMARY_HOVER = "#163F75"
TEXT_DARK = "#1F2A44"
TEXT_MUTED = "#6B7280"
REQUIRED_RED = "#E0434B"
ICON_TEAL = "#16A085"
SUCCESS_GREEN = "#1E8E5A"
WARNING_AMBER = "#B8760B"
ERROR_RED = "#C0392B"
STEP_INACTIVE_BORDER = "#C7CFDC"


def _load_json(path: Path) -> dict:
    if not path.exists():
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return {}


def _save_json(path: Path, data: dict) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, sort_keys=True)


def load_category_overrides() -> dict[str, str]:
    return {str(k): str(v) for k, v in _load_json(CATEGORY_OVERRIDES_PATH).items()}


def save_category_overrides(overrides: dict[str, str]) -> None:
    _save_json(CATEGORY_OVERRIDES_PATH, overrides)


def load_column_overrides() -> dict[str, dict[str, str]]:
    data = _load_json(COLUMN_OVERRIDES_PATH)
    return {str(k): {str(f): str(h) for f, h in v.items()} for k, v in data.items()}


def save_column_overrides(overrides: dict[str, dict[str, str]]) -> None:
    _save_json(COLUMN_OVERRIDES_PATH, overrides)


def _font(size: int, weight: str = "normal", slant: str = "roman") -> ctk.CTkFont:
    return ctk.CTkFont(family="Segoe UI", size=size, weight=weight, slant=slant)


class App:
    def __init__(self, root: ctk.CTk) -> None:
        self.root = root
        root.title(APP_TITLE)
        root.geometry("980x680")
        root.minsize(860, 600)
        root.configure(fg_color=BG)

        self.op_path = StringVar()
        self.ip_path = StringVar()
        self.previous_path = StringVar()
        self.output_dir = StringVar(value=str(REPORTS_DIR))
        self.status = StringVar(value="Ready.")

        self._pending_result: BuildResult | None = None
        self._pending_error: MissingColumnsError | None = None

        self._build_chrome()
        self._build_step1()
        self._build_step2()
        self.step3 = ctk.CTkFrame(self.content, fg_color="transparent")

        self._show_step(1)

    # -- shell: title bar substitute + stepper -----------------------------

    def _build_chrome(self) -> None:
        header = ctk.CTkFrame(self.root, fg_color=CARD_BG, corner_radius=0, height=56)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)
        ctk.CTkLabel(header, text="\U0001F4CA", font=_font(20), text_color=PRIMARY).pack(
            side="left", padx=(20, 8), pady=10
        )
        ctk.CTkLabel(header, text=APP_TITLE, font=_font(17, "bold"), text_color=TEXT_DARK).pack(
            side="left", pady=10
        )

        stepper_bar = ctk.CTkFrame(self.root, fg_color=BG, corner_radius=0)
        stepper_bar.pack(fill="x", side="top", padx=24, pady=(18, 6))
        self._build_stepper(stepper_bar)

        self.content = ctk.CTkFrame(self.root, fg_color=BG, corner_radius=0)
        self.content.pack(fill="both", expand=True, padx=24, pady=(6, 20))

    def _build_stepper(self, parent: ctk.CTkFrame) -> None:
        self._step_circles: list[ctk.CTkLabel] = []
        self._step_labels: list[ctk.CTkLabel] = []
        self._step_lines: list[ctk.CTkFrame] = []
        names = ["Add reports", "Generate", "Review results"]
        # Each step takes 2 columns (circle, label) plus 1 more for the
        # connector line after it (except the last step). Using `i * 2` for
        # the column index was wrong -- it only advances 2 columns per step,
        # so a step's circle landed in the same column as the previous
        # step's line. A running counter avoids that collision.
        col = 0
        for i, name in enumerate(names):
            circle = ctk.CTkLabel(
                parent, text=str(i + 1), width=32, height=32, corner_radius=16,
                fg_color="white", text_color=TEXT_MUTED, font=_font(13, "bold"), anchor="center",
            )
            circle.grid(row=0, column=col, padx=(0, 10))
            col += 1
            label = ctk.CTkLabel(parent, text=name, font=_font(13, "bold"), text_color=TEXT_MUTED, anchor="w")
            label.grid(row=0, column=col, sticky="w")
            parent.grid_columnconfigure(col, weight=0)
            col += 1
            self._step_circles.append(circle)
            self._step_labels.append(label)
            if i < len(names) - 1:
                line = ctk.CTkFrame(parent, fg_color=STEP_INACTIVE_BORDER, height=2)
                line.grid(row=0, column=col, sticky="ew", padx=16)
                parent.grid_columnconfigure(col, weight=1)
                col += 1
                self._step_lines.append(line)

    def _show_step(self, n: int) -> None:
        for i, (circle, label) in enumerate(zip(self._step_circles, self._step_labels), start=1):
            if i < n:
                circle.configure(fg_color=PRIMARY, text_color="white", text="✓")
                label.configure(text_color=PRIMARY)
            elif i == n:
                circle.configure(fg_color=PRIMARY, text_color="white", text=str(i))
                label.configure(text_color=TEXT_DARK)
            else:
                circle.configure(fg_color="white", text_color=TEXT_MUTED, text=str(i))
                label.configure(text_color=TEXT_MUTED)
        for i, line in enumerate(self._step_lines, start=1):
            line.configure(fg_color=PRIMARY if i < n else STEP_INACTIVE_BORDER)

        for page in (self.step1, self.step2, self.step3):
            page.pack_forget()
        {1: self.step1, 2: self.step2, 3: self.step3}[n].pack(fill="both", expand=True)

    # -- card helper ---------------------------------------------------------

    def _card(self, parent) -> ctk.CTkFrame:
        card = ctk.CTkFrame(parent, fg_color=CARD_BG, corner_radius=12, border_width=1, border_color=BORDER)
        return card

    # -- Step 1: Add reports -------------------------------------------------

    def _build_step1(self) -> None:
        self.step1 = ctk.CTkFrame(self.content, fg_color="transparent")

        card = self._card(self.step1)
        card.pack(fill="both", expand=True)

        head = ctk.CTkFrame(card, fg_color="transparent")
        head.pack(fill="x", padx=24, pady=(20, 10))
        ctk.CTkLabel(head, text="\U0001F4C4", font=_font(22), text_color=ICON_TEAL).grid(
            row=0, column=0, rowspan=2, sticky="n", padx=(0, 12)
        )
        ctk.CTkLabel(head, text="Input files", font=_font(18, "bold"), text_color=TEXT_DARK).grid(
            row=0, column=1, sticky="w"
        )
        ctk.CTkLabel(
            head, text="Choose the reports needed to calculate AR outstanding.",
            font=_font(13), text_color=TEXT_MUTED,
        ).grid(row=1, column=1, sticky="w")

        self._file_row(
            card, "Total Bill Wise Report (OP)", required=True, var=self.op_path,
            hint="Excel or CSV files (.xlsx, .xls, .csv)",
            filetypes=[("Excel or CSV files", "*.xlsx *.xls *.csv"), ("All files", "*.*")],
        )
        self._file_row(
            card, "Discharge Wise IP Revenue Report", required=True, var=self.ip_path,
            hint="Excel or CSV files (.xlsx, .xls, .csv)",
            filetypes=[("Excel or CSV files", "*.xlsx *.xls *.csv"), ("All files", "*.*")],
        )
        self._file_row(
            card, "Previous AR Report", required=False, var=self.previous_path,
            hint="Excel files (.xlsx, .xls)",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )

        out_row = ctk.CTkFrame(card, fg_color="transparent")
        out_row.pack(fill="x", padx=24, pady=(4, 18))
        ctk.CTkLabel(out_row, text="\U0001F4C1", font=_font(16), text_color=ICON_TEAL).grid(
            row=0, column=0, rowspan=2, padx=(0, 10), sticky="n"
        )
        ctk.CTkLabel(out_row, text="Reports will be saved to:", font=_font(13, "bold"), text_color=TEXT_DARK).grid(
            row=0, column=1, sticky="w"
        )
        self.output_dir_label = ctk.CTkLabel(
            out_row, textvariable=self.output_dir, font=_font(12), text_color=TEXT_MUTED, anchor="w"
        )
        self.output_dir_label.grid(row=1, column=1, sticky="w")
        out_row.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(
            out_row, text="✏ Change folder", fg_color="transparent", hover_color="#F0F4FA",
            text_color=PRIMARY, font=_font(13, "bold"), command=self._change_output_dir,
        ).grid(row=0, column=2, rowspan=2, sticky="e")

        btn_row = ctk.CTkFrame(self.step1, fg_color="transparent")
        btn_row.pack(fill="x", pady=(16, 0))
        self.generate_btn = ctk.CTkButton(
            btn_row, text="\U0001F4CA  Generate AR Report", font=_font(15, "bold"),
            fg_color=PRIMARY, hover_color=PRIMARY_HOVER, height=48, corner_radius=8,
            command=self._on_generate,
        )
        self.generate_btn.pack(side="right")

    LEFT_COL_WIDTH = 230

    def _file_row(self, card, title: str, required: bool, var: StringVar, hint: str, filetypes: list) -> None:
        row = ctk.CTkFrame(card, fg_color="transparent")
        row.pack(fill="x", padx=24, pady=8)
        # A fixed minsize on the grid COLUMN, not a width-locked child frame,
        # is what actually keeps the file boxes aligned across all three rows
        # -- each row is an independent grid, so without this, a longer title
        # (e.g. "Discharge Wise IP Revenue Report") makes column 0 wider only
        # for that row, shifting its box out of line with the others.
        row.grid_columnconfigure(0, minsize=self.LEFT_COL_WIDTH)
        row.grid_columnconfigure(1, weight=1)

        left = ctk.CTkFrame(row, fg_color="transparent")
        left.grid(row=0, column=0, sticky="nw", padx=(0, 16))
        title_line = ctk.CTkFrame(left, fg_color="transparent")
        title_line.pack(anchor="w")
        ctk.CTkLabel(
            title_line, text=title, font=_font(14, "bold"), text_color=TEXT_DARK,
            anchor="w", justify="left", wraplength=self.LEFT_COL_WIDTH - 20,
        ).pack(side="left")
        if required:
            ctk.CTkLabel(title_line, text=" *", font=_font(14, "bold"), text_color=REQUIRED_RED).pack(side="left")
        else:
            ctk.CTkLabel(title_line, text="  (optional)", font=_font(12), text_color=TEXT_MUTED).pack(side="left")
        ctk.CTkLabel(
            left, text="Required" if required else "Optional", font=_font(12), text_color=TEXT_MUTED,
        ).pack(anchor="w")

        box = ctk.CTkFrame(row, fg_color="#FBFCFE", corner_radius=8, border_width=1, border_color=FIELD_BORDER)
        box.grid(row=0, column=1, sticky="nsew")
        inner = ctk.CTkFrame(box, fg_color="transparent")
        inner.pack(fill="x", padx=14, pady=10)
        ctk.CTkLabel(inner, text="\U0001F4C4", font=_font(16), text_color=TEXT_MUTED).pack(side="left", padx=(0, 10))
        text_col = ctk.CTkFrame(inner, fg_color="transparent")
        text_col.pack(side="left", fill="x", expand=True)
        name_label = ctk.CTkLabel(text_col, text="No file selected", font=_font(13), text_color=TEXT_MUTED, anchor="w")
        name_label.pack(fill="x", anchor="w")
        ctk.CTkLabel(text_col, text=hint, font=_font(11), text_color=TEXT_MUTED, anchor="w").pack(fill="x", anchor="w")

        def refresh(*_a):
            path = var.get()
            name_label.configure(
                text=Path(path).name if path else "No file selected",
                text_color=TEXT_DARK if path else TEXT_MUTED,
            )
        var.trace_add("write", refresh)

        ctk.CTkButton(
            row, text="\U0001F4C1 Browse...", width=120, height=34, corner_radius=6,
            fg_color="white", hover_color="#F0F4FA", text_color=PRIMARY,
            border_width=1, border_color=PRIMARY, font=_font(13),
            command=lambda: self._pick_file(var, filetypes),
        ).grid(row=0, column=2, sticky="e", padx=(16, 0))

    def _pick_file(self, var: StringVar, filetypes: list) -> None:
        path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            var.set(path)

    def _change_output_dir(self) -> None:
        path = filedialog.askdirectory(initialdir=self.output_dir.get() or str(APP_DIR))
        if path:
            self.output_dir.set(path)

    # -- Step 2: Generate (progress) -----------------------------------------

    def _build_step2(self) -> None:
        self.step2 = ctk.CTkFrame(self.content, fg_color="transparent")
        card = self._card(self.step2)
        card.pack(fill="both", expand=True)
        # pack(expand=True) centers regardless of window size/resizing; a
        # guessed relx/rely fraction drifts off-center when the card's actual
        # height doesn't match what the fraction assumed.
        center = ctk.CTkFrame(card, fg_color="transparent")
        center.pack(expand=True)
        ctk.CTkLabel(center, text="⏳", font=_font(40)).pack(pady=(0, 16))
        ctk.CTkLabel(center, text="Generating your report...", font=_font(17, "bold"), text_color=TEXT_DARK).pack()
        self.progress = ctk.CTkProgressBar(
            center, width=380, height=10, corner_radius=5, mode="indeterminate",
            fg_color="#E3E8F0", progress_color=PRIMARY,
        )
        self.progress.pack(pady=20)
        ctk.CTkLabel(center, textvariable=self.status, font=_font(12), text_color=TEXT_MUTED, wraplength=500).pack()

    # -- generate flow --------------------------------------------------------

    def _on_generate(self) -> None:
        if not self.op_path.get() or not self.ip_path.get():
            messagebox.showerror(APP_TITLE, "Please select the OP and IP report files.")
            return
        self._show_step(2)
        self.status.set("Working...")
        self.progress.start()
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self) -> None:
        """Always re-reads both override files from disk, so every retry
        (after a NEEDS REVIEW decision or a column mapping) picks up
        whatever was just saved."""
        category_overrides = load_category_overrides()
        column_overrides = load_column_overrides()
        try:
            result = build_report(
                self.op_path.get(), self.ip_path.get(), self.previous_path.get() or None,
                output_dir=self.output_dir.get() or None, category_overrides=category_overrides,
                op_column_overrides=column_overrides.get("OP"),
                ip_column_overrides=column_overrides.get("IP"),
            )
            self.root.after(0, self._handle_result, result)
        except MissingColumnsError as e:
            self.root.after(0, self._render_missing_columns, e)
        except SourceFormatError as e:
            self.root.after(0, self._render_failure, str(e))
        except Exception:
            self.root.after(0, self._render_failure, traceback.format_exc())

    def _handle_result(self, result: BuildResult) -> None:
        if not result.needs_review.empty:
            self._render_needs_review(result)
        else:
            self._render_success(result)

    # -- Step 3 content builders ----------------------------------------------

    def _reset_step3(self) -> ctk.CTkFrame:
        self.progress.stop()
        for child in self.step3.winfo_children():
            child.destroy()
        self._show_step(3)
        return self.step3

    def _result_heading(self, parent, icon: str, title: str, color: str, subtitle: str = "") -> ctk.CTkFrame:
        card = self._card(parent)
        card.pack(fill="x", pady=(0, 14))
        row = ctk.CTkFrame(card, fg_color="transparent")
        row.pack(fill="x", padx=24, pady=18)
        ctk.CTkLabel(row, text=icon, font=_font(22)).pack(side="left", padx=(0, 10))
        text_col = ctk.CTkFrame(row, fg_color="transparent")
        text_col.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(text_col, text=title, font=_font(17, "bold"), text_color=color, anchor="w").pack(fill="x")
        if subtitle:
            ctk.CTkLabel(
                text_col, text=subtitle, font=_font(13), text_color=TEXT_MUTED, anchor="w", justify="left",
                wraplength=760,
            ).pack(fill="x")
        return card

    def _render_success(self, result: BuildResult) -> None:
        page = self._reset_step3()
        self._result_heading(page, "✅", "Report generated", SUCCESS_GREEN, str(result.output_path))

        card = self._card(page)
        card.pack(fill="x", pady=(0, 14))
        body = ctk.CTkFrame(card, fg_color="transparent")
        body.pack(fill="x", padx=24, pady=18)
        rows = [
            ("Report date", f"{result.report_date:%d-%m-%Y}"),
            ("Public Ins", f"{result.bill_count_public} bill(s), total {result.public_total:,.2f}"),
            ("TPA & CPR", f"{result.bill_count_tpa} bill(s), total {result.tpa_total:,.2f}"),
            ("OP GNL", f"{result.bill_count_gnl} bill(s)"),
        ]
        for i, (label, value) in enumerate(rows):
            ctk.CTkLabel(body, text=label, font=_font(13, "bold"), text_color=TEXT_DARK, anchor="w", width=140).grid(
                row=i, column=0, sticky="w", pady=3
            )
            ctk.CTkLabel(body, text=value, font=_font(13), text_color=TEXT_DARK, anchor="w").grid(
                row=i, column=1, sticky="w", pady=3
            )
        if result.warnings:
            notes = ctk.CTkFrame(card, fg_color="#F7F9FC", corner_radius=8)
            notes.pack(fill="x", padx=24, pady=(0, 18))
            for w in result.warnings:
                ctk.CTkLabel(
                    notes, text=f"• {w}", font=_font(12), text_color=TEXT_MUTED, anchor="w",
                    justify="left", wraplength=760,
                ).pack(fill="x", padx=12, pady=4)

        btn_row = ctk.CTkFrame(page, fg_color="transparent")
        btn_row.pack(fill="x")
        ctk.CTkButton(
            btn_row, text="Open output folder", fg_color="white", hover_color="#F0F4FA",
            text_color=PRIMARY, border_width=1, border_color=PRIMARY, font=_font(13, "bold"),
            command=lambda: os.startfile(result.output_path.parent),  # noqa: S606 -- Windows-only app
        ).pack(side="left")
        ctk.CTkButton(
            btn_row, text="Start New Report", fg_color=PRIMARY, hover_color=PRIMARY_HOVER,
            font=_font(13, "bold"), command=lambda: self._show_step(1),
        ).pack(side="right")

    def _render_needs_review(self, result: BuildResult) -> None:
        self._pending_result = result
        page = self._reset_step3()
        df = result.needs_review.copy()
        df["OverrideKey"] = [override_key(c, n) for c, n in zip(df["Category"], df["CompanyName"])]
        n_groups = df["OverrideKey"].nunique()

        self._result_heading(
            page, "⚠", "Needs review -- unrecognised categories", WARNING_AMBER,
            f"{len(df)} bill(s) across {n_groups} group(s) were routed to TPA & CPR under a category this tool "
            "doesn't recognise. They're grouped by Category + Company Name, since bills can share the same "
            "category text but belong to different payers needing different decisions. Pick where each group "
            "actually belongs, then Apply & Regenerate. Decisions are remembered for future reports too.",
        )

        groups = sorted(df.groupby("OverrideKey", sort=False), key=lambda kv: kv[1]["PaymentDue"].sum(), reverse=True)

        scroll = ctk.CTkScrollableFrame(page, fg_color=CARD_BG, corner_radius=12, border_width=1, border_color=BORDER,
                                         height=320)
        scroll.pack(fill="both", expand=True, pady=(0, 14))

        col_widths = [20, 12, 26, 14]
        decision_vars: dict[str, StringVar] = {}
        for key, group in groups:
            sample_category = group["Category"].iloc[0]
            sample_company = group["CompanyName"].iloc[0]
            total_due = group["PaymentDue"].sum()
            header_row = ctk.CTkFrame(scroll, fg_color="transparent")
            header_row.pack(fill="x", padx=12, pady=(10, 2))
            ctk.CTkLabel(
                header_row, text=f"{sample_category}  /  {sample_company}  --  {len(group)} bill(s), total {total_due:,.2f}",
                font=_font(13, "bold"), text_color=TEXT_DARK, anchor="w",
            ).pack(side="left", fill="x", expand=True)
            var = StringVar(value="Keep as TPA & CPR")
            ctk.CTkComboBox(
                header_row, variable=var, values=list(BUCKET_CHOICES.keys()), width=200, state="readonly",
            ).pack(side="right")
            decision_vars[key] = var

            detail_header = ctk.CTkFrame(scroll, fg_color="transparent")
            detail_header.pack(fill="x", padx=12)
            for text, w in zip(["Bill No.", "Date", "Company Name", "Payment Due"], col_widths):
                ctk.CTkLabel(detail_header, text=text, font=_font(11, slant="italic"), text_color=TEXT_MUTED,
                             width=w * 7, anchor="w").pack(side="left")

            for _, bill in group.iterrows():
                date_val = bill["Date"]
                date_str = date_val.strftime("%d-%m-%Y") if pd.notna(date_val) else ""
                detail_row = ctk.CTkFrame(scroll, fg_color="transparent")
                detail_row.pack(fill="x", padx=12)
                for text, w in zip(
                    [str(bill["BillNo"]), date_str, str(bill["CompanyName"]), f"{bill['PaymentDue']:,.2f}"],
                    col_widths,
                ):
                    ctk.CTkLabel(detail_row, text=text, font=_font(12), text_color=TEXT_DARK, width=w * 7,
                                 anchor="w").pack(side="left")

        btn_row = ctk.CTkFrame(page, fg_color="transparent")
        btn_row.pack(fill="x")
        ctk.CTkButton(
            btn_row, text="Apply & Regenerate", fg_color=PRIMARY, hover_color=PRIMARY_HOVER,
            font=_font(13, "bold"), command=lambda: self._apply_review(decision_vars),
        ).pack(side="left")
        ctk.CTkButton(
            btn_row, text="Skip (keep this report as generated)", fg_color="white", hover_color="#F0F4FA",
            text_color=PRIMARY, border_width=1, border_color=PRIMARY, font=_font(13, "bold"),
            command=lambda: self._render_success(result),
        ).pack(side="left", padx=8)

    def _apply_review(self, decision_vars: dict[str, StringVar]) -> None:
        new_overrides = {key: BUCKET_CHOICES[var.get()] for key, var in decision_vars.items()}
        merged = load_category_overrides()
        merged.update(new_overrides)
        save_category_overrides(merged)
        self._pending_result = None
        self._show_step(2)
        self.status.set("Applying your decisions and regenerating the report...")
        self.progress.start()
        threading.Thread(target=self._run, daemon=True).start()

    def _render_missing_columns(self, error: MissingColumnsError) -> None:
        self._pending_error = error
        page = self._reset_step3()
        self._result_heading(
            page, "⚠", "New column layout -- map missing fields", WARNING_AMBER,
            f"{error.report_label} doesn't have a column this tool recognises for the field(s) below -- the "
            "report's layout has probably changed. Pick which of the file's actual columns each one really is, "
            "then Apply & Retry. Mappings are remembered for future reports too.",
        )

        card = self._card(page)
        card.pack(fill="both", expand=True, pady=(0, 14))
        body = ctk.CTkFrame(card, fg_color="transparent")
        body.pack(fill="x", padx=24, pady=18)

        display_headers = [str(h).replace("\n", " ") for h in error.available_headers]
        field_vars: dict[str, StringVar] = {}
        for i, field in enumerate(error.missing_fields):
            ctk.CTkLabel(body, text=FIELD_LABELS.get(field, field), font=_font(13, "bold"), text_color=TEXT_DARK,
                         width=240, anchor="w").grid(row=i, column=0, sticky="w", pady=4)
            var = StringVar(value="-- not mapped --")
            ctk.CTkComboBox(
                body, variable=var, values=["-- not mapped --"] + display_headers, width=320, state="readonly",
            ).grid(row=i, column=1, sticky="w", pady=4)
            field_vars[field] = var

        ctk.CTkLabel(
            card, text=f"This file's actual headers: {', '.join(display_headers)}", font=_font(11),
            text_color=TEXT_MUTED, anchor="w", justify="left", wraplength=800,
        ).pack(fill="x", padx=24, pady=(0, 16))

        btn_row = ctk.CTkFrame(page, fg_color="transparent")
        btn_row.pack(fill="x")
        ctk.CTkButton(
            btn_row, text="Apply & Retry", fg_color=PRIMARY, hover_color=PRIMARY_HOVER, font=_font(13, "bold"),
            command=lambda: self._apply_column_mapping(error, field_vars),
        ).pack(side="left")
        ctk.CTkButton(
            btn_row, text="Cancel", fg_color="white", hover_color="#F0F4FA", text_color=PRIMARY,
            border_width=1, border_color=PRIMARY, font=_font(13, "bold"),
            command=lambda: self._render_failure(str(error)),
        ).pack(side="left", padx=8)

    def _apply_column_mapping(self, error: MissingColumnsError, field_vars: dict[str, StringVar]) -> None:
        chosen = {field: var.get() for field, var in field_vars.items() if var.get() != "-- not mapped --"}
        if not chosen:
            messagebox.showerror(APP_TITLE, "Map at least one field before applying, or Cancel.")
            return
        report_key = REPORT_LABEL_TO_KEY.get(error.report_label, error.report_label)
        all_overrides = load_column_overrides()
        report_overrides = all_overrides.setdefault(report_key, {})
        report_overrides.update(chosen)
        save_column_overrides(all_overrides)
        self._pending_error = None
        self._show_step(2)
        self.status.set("Applying your column mapping and retrying...")
        self.progress.start()
        threading.Thread(target=self._run, daemon=True).start()

    def _render_failure(self, message: str) -> None:
        page = self._reset_step3()
        self._result_heading(page, "❌", "Something went wrong", ERROR_RED)

        card = self._card(page)
        card.pack(fill="both", expand=True, pady=(0, 14))
        box = ctk.CTkTextbox(card, fg_color="#FBFCFE", font=_font(12), text_color=TEXT_DARK, wrap="word")
        box.pack(fill="both", expand=True, padx=16, pady=16)
        box.insert("1.0", message)
        box.configure(state="disabled")

        btn_row = ctk.CTkFrame(page, fg_color="transparent")
        btn_row.pack(fill="x")
        ctk.CTkButton(
            btn_row, text="Back to Step 1", fg_color=PRIMARY, hover_color=PRIMARY_HOVER, font=_font(13, "bold"),
            command=lambda: self._show_step(1),
        ).pack(side="left")


def main() -> None:
    ctk.set_appearance_mode("light")
    root = ctk.CTk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
