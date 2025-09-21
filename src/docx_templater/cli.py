from practice2.src.docx_templater.helpers import replace_placeholders, extract_csv_data
from practice2.src.docx_templater.settings import TEST_DOC_PATH, OUTPUT_DIR

if __name__ == "__main__":
    items = extract_csv_data()
    for n, sample in enumerate(items):
        replace_placeholders(sample, docx_path=TEST_DOC_PATH, out_path=f"{OUTPUT_DIR}/document_{n}.docx")