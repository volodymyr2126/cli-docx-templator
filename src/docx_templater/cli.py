import argparse


from practice2.src.docx_templater.helpers import dialogue
from practice2.src.docx_templater.settings import DEFAULT_DATA_PATH, DEFAULT_DOC_PATH, OUTPUT_DIR, DEFAULT_PATTERN

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate DOCX from template and CSV data")

    parser.add_argument("--template", required=False, default=DEFAULT_DOC_PATH, help="Path to DOCX template")
    parser.add_argument("--csv", required=False, default=DEFAULT_DATA_PATH, help="Output directory or file prefix")
    parser.add_argument("--outdir", required=False, default=OUTPUT_DIR, help="Path to CSV data")
    parser.add_argument("--filename_pattern", required=False, default=DEFAULT_PATTERN, help="Path to CSV data")

    args = parser.parse_args()

    dialogue(
        docx_path=args.template,
        out_path=args.outdir,
        data_path=args.csv,
        file_name_pattern=args.filename_pattern
    )
