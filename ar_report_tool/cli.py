"""Command-line entry point, mainly for testing logic.py without the GUI.

Usage:
    python cli.py <op_file> <ip_file> [previous_report_file] [output_dir]

The report date is inferred from the data. [previous_report_file] is
optional -- when given, it supplies Table 2's (Decrease) previous-report
column, taken from that file's own Table 1 old-bucket figures for
Public/TPA. Without it, the output has only Table 1 (current outstanding
totals). Pass "-" in its place if you need to supply [output_dir] without a
previous report.
"""
import sys

from logic import build_report


def main() -> None:
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    op_path, ip_path = sys.argv[1], sys.argv[2]
    previous_arg = sys.argv[3] if len(sys.argv) > 3 else None
    previous_path = None if previous_arg in (None, "-") else previous_arg
    output_dir = sys.argv[4] if len(sys.argv) > 4 else None

    result = build_report(op_path, ip_path, previous_path, output_dir)

    print(f"Output: {result.output_path}")
    print(f"Report date: {result.report_date:%d-%m-%Y}")
    print(f"Public Ins: {result.bill_count_public} bills, total {result.public_total:,.2f}")
    print(f"TPA & CPR : {result.bill_count_tpa} bills, total {result.tpa_total:,.2f}")
    print(f"OP GNL    : {result.bill_count_gnl} bills")
    if result.previous_report_date:
        print(f"Previous report date used for Table 2: {result.previous_report_date}")
    if not result.needs_review.empty:
        print(f"Needs review: {len(result.needs_review)} bill(s) -- see NEEDS REVIEW sheet")
    for w in result.warnings:
        print(f"NOTE: {w}")


if __name__ == "__main__":
    main()
