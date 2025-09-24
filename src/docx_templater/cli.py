import argparse
from pathlib import Path
from helpers import dialogue
from settings import DEFAULT_DATA_PATH, DEFAULT_DOC_PATH, OUTPUT_DIR, DEFAULT_PATTERN

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate DOCX from template and CSV data")

    parser.add_argument("--template", required=True, default=DEFAULT_DOC_PATH,
                        help="Path to DOCX template")
    parser.add_argument("--csv", required=True, default=DEFAULT_DATA_PATH,
                        help="Output directory or file prefix")
    parser.add_argument("--outdir", required=True, default=OUTPUT_DIR,
                        help="Path to CSV data")
    parser.add_argument("--filename-pattern", required=True, default=DEFAULT_PATTERN,
                        help="Filename pattern of output files")
    parser.add_argument("--dry-run", action="store_true", default=False,
                        help="Run without writing any DOCX files (just show what would happen)")
    parser.add_argument("--single", type=int, default=None, metavar="N",
                        help="Process only a single row from the CSV (by index, 0-based)")
    args = parser.parse_args()

    if not args.dry_run:
        out_dir = Path(args.outdir)
        out_dir.mkdir(parents=True, exist_ok=True)

    dialogue(
        docx_path=args.template,
        out_path=args.outdir,
        data_path=args.csv,
        file_name_pattern=args.filename_pattern,
        dry_run=args.dry_run,
        single=args.single
    )
