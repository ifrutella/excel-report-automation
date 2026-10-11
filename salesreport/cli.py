"""Command line entry point: python -m salesreport <folder> [-o report.xlsx]"""

import argparse
import sys
from pathlib import Path

from .clean import clean, load_folder
from .excel import write_report


def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="salesreport",
        description="Combine messy order exports (CSV/Excel) into one clean, formatted Excel report.",
    )
    parser.add_argument("folder", help="folder with the exported .csv / .xlsx files")
    parser.add_argument("-o", "--output", default="sales_report.xlsx", help="where to save the report (default: sales_report.xlsx)")
    parser.add_argument("--title", default="Sales report", help="title on the Summary sheet")
    args = parser.parse_args(argv)

    try:
        frames, files = load_folder(args.folder)
    except (FileNotFoundError, ValueError) as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1

    rows_read = sum(len(f) for f in frames)
    data, issues = clean(frames)

    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    try:
        write_report(data, issues, files, output, title=args.title)
    except PermissionError:
        print(f"Error: cannot write {output}. Is the file open in Excel?", file=sys.stderr)
        return 1

    print(f"Read {rows_read} rows from {len(files)} files: {', '.join(files)}")
    print(f"Used {len(data)} rows, set aside {len(issues)} (reasons are on the Issues sheet)")
    print(f"Revenue {data['revenue'].sum():,.0f} kr from {data['order_id'].nunique()} orders".replace(",", " "))
    print(f"Saved {output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
