# DOCX Templater CLI

A command-line tool to generate `.docx` documents from a template and CSV data.  
It replaces placeholders in a Word document (e.g., `{name}`, `{phone}`) with values from a CSV file and produces one or multiple output documents.

---

## Features

- Extracts all placeholders from a `.docx` template.
- Reads data from a CSV file (columns must match placeholder names or be mapped interactively).
- Automatically replaces placeholders with CSV values.
- Supports batch processing: creates one output document per CSV row.
- Allows customizing output filenames using patterns (e.g., `output_{name}_{id}.docx`).

---

## Installation

1. Clone the repository:

```bash
git clone https://github.com/volodymyr2126/cli-docx-templator.git
cd practice2
```
Install dependencies:
```bash
pip install -r requirements.txt
```

Required packages: lxml

## Usage
```bashpython src/docx_templater/cli.py \
    --template path/to/template.docx \
    --csv path/to/data.csv \
    --outdir output_directory \
    --filename_pattern "output_{name}_{id}.docx"
```

## CLI Arguments
| Argument                   | Description                                                                | Default                      |
|----------------------------| -------------------------------------------------------------------------- | ---------------------------- |
| `--template`               | Path to the DOCX template with placeholders                                | `settings.DEFAULT_DOC_PATH`  |
| `--csv`                    | Path to CSV file containing data for placeholders                          | `settings.DEFAULT_DATA_PATH` |
| `--outdir`                 | Directory to save generated DOCX files                                     | `settings.OUTPUT_DIR`        |
| `--filename_pattern`       | Pattern for naming output files. Use `{column_name}` placeholders from CSV | `settings.DEFAULT_PATTERN`   |


## Example

Suppose your template template.docx contains placeholders:

```
Name: {name}
Phone: {phone}
Department: {department}
```



CSV data.csv:

```
name,phone,department
John,4452384,Compliance
Steve,4456890,Marketing
```

Running:

```bash
python src/docx_templater/cli.py \
    --template template.docx \
    --csv data.csv \
    --outdir output \
    --filename_pattern "doc_{name}.docx"
```


Will generate:

```
output/doc_John.docx
output/doc_Steve.docx
```


Each document will have placeholders replaced with the corresponding CSV row values.

## Notes

- If the CSV columns do not match the template placeholders, the tool will prompt you to map missing variables interactively.

- Placeholders in the template should be enclosed in curly braces, e.g., {placeholder_name}.

- Sample data from CSV is displayed to help map columns correctly.

- Output filenames will replace spaces with underscores to ensure valid file names.