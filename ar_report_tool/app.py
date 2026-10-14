"""AR Outstanding Report -- desktop app.

Pick the two system report downloads plus the previous AR report, then click
Generate. The report date is inferred from the data and the output is always
saved into a "Reports" folder next to the .exe, so every generated report
collects in one place instead of scattering into wherever the OP file lives.
The previous AR report supplies Table 2's (Decrease) previous-report column.
This is the file PyInstaller packages into the .exe -- see BUILD.md /
README.md.
"""
from __future__ import annotations

import os
import sys
import threading
import traceback
from pathlib import Path
from tkinter import Tk, StringVar, filedialog, messagebox
from tkinter import ttk

from logic import SourceFormatError, build_report

APP_TITLE = "AR Outstanding Report Builder"

APP_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
REPORTS_DIR = APP_DIR / "Reports"


class App:
    def __init__(self, root: Tk) -> None:
        self.root = root
        root.title(APP_TITLE)
        root.geometry("640x320")
        root.resizable(False, False)

        self.op_path = StringVar()
        self.ip_path = StringVar()
        self.previous_path = StringVar()
        self.status = StringVar(value="Ready.")

        pad = {"padx": 10, "pady": 8}

        self._file_row(0, "Total Bill Wise Report (OP):", self.op_path, pad)
        self._file_row(1, "Discharge Wise IP Revenue Report:", self.ip_path, pad)
        self._file_row(2, "Previous AR Report (optional):", self.previous_path, pad)
        ttk.Label(
            root, text=f"Reports are saved to: {REPORTS_DIR}", foreground="gray",
        ).grid(row=3, column=0, columnspan=3, sticky="w", padx=10)

        self.generate_btn = ttk.Button(root, text="Generate AR Report", command=self._on_generate)
        self.generate_btn.grid(row=4, column=0, columnspan=3, pady=18)

        ttk.Label(root, textvariable=self.status, wraplength=600, justify="left").grid(
            row=5, column=0, columnspan=3, sticky="w", **pad
        )

    def _file_row(self, row: int, label: str, var: StringVar, pad: dict) -> None:
        ttk.Label(self.root, text=label).grid(row=row, column=0, sticky="w", **pad)
        ttk.Entry(self.root, textvariable=var, width=48).grid(row=row, column=1, **pad)
        ttk.Button(self.root, text="Browse...", command=lambda v=var: self._pick_file(v)).grid(
            row=row, column=2, **pad
        )

    def _pick_file(self, var: StringVar) -> None:
        path = filedialog.askopenfilename(
            filetypes=[("Excel or CSV files", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
        )
        if path:
            var.set(path)

    def _on_generate(self) -> None:
        if not self.op_path.get() or not self.ip_path.get():
            messagebox.showerror(APP_TITLE, "Please select the OP and IP report files.")
            return
        self.generate_btn.state(["disabled"])
        self.status.set("Working...")
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self) -> None:
        try:
            result = build_report(
                self.op_path.get(), self.ip_path.get(), self.previous_path.get() or None,
                output_dir=REPORTS_DIR,
            )
            lines = [
                f"Done: {result.output_path}",
                f"Report date (from the data): {result.report_date:%d-%m-%Y}",
                f"Public Ins: {result.bill_count_public} bills, total {result.public_total:,.2f}",
                f"TPA & CPR: {result.bill_count_tpa} bills, total {result.tpa_total:,.2f}",
                f"OP GNL: {result.bill_count_gnl} bills",
            ]
            if not result.needs_review.empty:
                lines.append(f"{len(result.needs_review)} bill(s) flagged on the NEEDS REVIEW sheet -- please check.")
            lines.extend(f"Note: {w}" for w in result.warnings)
            self.root.after(0, self._done, "\n".join(lines), result.output_path)
        except SourceFormatError as e:
            self.root.after(0, self._fail, str(e))
        except Exception:
            self.root.after(0, self._fail, traceback.format_exc())

    def _done(self, message: str, output_path: Path) -> None:
        self.status.set(message)
        self.generate_btn.state(["!disabled"])
        if messagebox.askyesno(APP_TITLE, "Report generated. Open the output folder now?"):
            os.startfile(output_path.parent)  # noqa: S606 -- Windows-only app

    def _fail(self, message: str) -> None:
        self.status.set("Failed -- see details below.\n" + message)
        self.generate_btn.state(["!disabled"])
        messagebox.showerror(APP_TITLE, message)


def main() -> None:
    root = Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
