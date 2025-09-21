import zipfile
import os
import shutil
from typing import List
import csv
from lxml import etree
from copy import deepcopy

from practice2.src.docx_templater.settings import NS, UNZIP_OUTPUT, OUTPUT_DIR


def process_paragraph(p: etree.Element, variables: dict) -> None:
    new_runs = []
    inside = False
    placeholder = ""
    first_run = None

    for run in p.findall("w:r", NS):
        texts = run.findall("w:t", NS)
        run_text = "".join([t.text for t in texts if t.text])

        for ch in run_text:
            if ch == "{":
                inside = True
                placeholder = ""
                first_run = run
            elif ch == "}":
                inside = False
                value = variables.get(placeholder, f"{{{placeholder}}}")
                new_run = deepcopy(first_run)
                for t in new_run.findall("w:t", NS):
                    t.text = value
                new_runs.append(new_run)
                first_run = None
            elif inside:
                placeholder += ch
            else:
                new_run = deepcopy(run)
                for t in new_run.findall("w:t", NS):
                    t.text = ch
                new_runs.append(new_run)

    for run in p.findall("w:r", NS):
        p.remove(run)
    for run in new_runs:
        p.append(run)


def replace_placeholders(variables: dict,
                         docx_path: str = "./samples/template_sample.docx",
                         out_path: str = "./output") -> None:
    with zipfile.ZipFile(docx_path, "r") as zin:
        zin.extractall(UNZIP_OUTPUT)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    xml_path = os.path.join(UNZIP_OUTPUT, "word", "document.xml")
    tree = etree.parse(xml_path)
    root = tree.getroot()

    for p in root.findall(".//w:p", NS):
        process_paragraph(p, variables)
    tree.write(xml_path, xml_declaration=True, encoding="UTF-8", standalone="yes")

    with zipfile.ZipFile(out_path, "w") as zout:
        for folder, _, files in os.walk(UNZIP_OUTPUT):
            for f in files:
                file_path = os.path.join(folder, f)
                arcname = os.path.relpath(file_path, UNZIP_OUTPUT)
                zout.write(file_path, arcname)
    shutil.rmtree(UNZIP_OUTPUT)

def extract_csv_data(path_to_csv: str = "./samples/test.csv") -> List[dict]:
    with open(path_to_csv, mode="r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        data = list(reader)
        return data