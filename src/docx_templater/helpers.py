import zipfile
import os
import shutil
from lxml import etree
from copy import deepcopy
import csv
from typing import List
import re

from practice2.src.docx_templater.settings import UNZIP_OUTPUT, OUTPUT_DIR, NS


def get_xml_from_docx(docx_path: str) -> etree.ElementTree:
    with zipfile.ZipFile(docx_path, "r") as zin:
        zin.extractall(UNZIP_OUTPUT)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    xml_path = os.path.join(UNZIP_OUTPUT, "word", "document.xml")
    tree = etree.parse(xml_path)
    return tree, xml_path


def extract_csv_data(path_to_csv: str = "./samples/test.csv") -> List[dict]:
    with open(path_to_csv, mode="r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        data = list(reader)
        return data


def get_all_text(tree: etree.Element) -> str:
    return "".join([node.text for node in tree.getroot().findall(".//w:t", NS)])


def get_all_variables(text: str) -> set:
    return set(re.findall(r"\{([^}]+)\}", text))


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


def replace_placeholders(tree: etree.ElementTree, xml_path: str, out_path: str, variables: dict) -> None:
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
    print(get_all_text(tree))


def get_missing_positions(old_csv_vars: list, new_csv_vars: list) -> list:
    missing_positions = []
    for n, var in enumerate(old_csv_vars):
        if var not in new_csv_vars:
            missing_positions.append(n)
    return missing_positions


def compare_vars(csv_vars: list, docx_vars: set, data: list) -> list:
    missing_vars = []
    new_csv_vars = []
    if len(csv_vars) != len(docx_vars):
        print(f"Number of csv columns doesn't match number of columns in docx file. Please edit your csv file.")
        return []
    for n, var in enumerate(docx_vars):
        if var not in csv_vars:
            missing_vars.append(var)
        else:
            new_csv_vars.append(var)
    if not missing_vars:
        return csv_vars
    print(f"You have {len(missing_vars)} unmapped columns.")

    missing_positions = get_missing_positions(old_csv_vars=csv_vars, new_csv_vars=new_csv_vars)

    for pos in missing_positions:
        print(f"Sample data for position {pos}: {','.join([item[csv_vars[pos]] for item in data[:5]])}")
    print(f"Please, map missing positions to missing vars")

    mapping = {}
    for var in missing_vars:
        while True:
            pos_str = input(f"Which position should '{var}' map to? (choose from {missing_positions}) ")
            try:
                pos = int(pos_str)
            except ValueError:
                print("❌ Please enter a number.")
                continue
            if pos in missing_positions:
                mapping[pos] = var
                missing_positions.remove(pos)
                break
            else:
                print("❌ Invalid position, try again.")

    final_csv_vars = csv_vars.copy()
    for pos, mapped_var in mapping.items():
        final_csv_vars[pos] = mapped_var

    return final_csv_vars


def dialogue(docx_path: str, out_path: str, data_path: str, file_name_pattern: str) -> None:
    tree, xml_path = get_xml_from_docx(docx_path=docx_path)
    items = extract_csv_data(path_to_csv=data_path)

    if not items:
        print(f"Your csv file {data_path} is empty. Please put some values there.")

    text = get_all_text(tree=tree)
    if not text:
        print(f"Your word doc {docx_path} is empty. Please put some templated text there.")

    var_names = get_all_variables(text=text)
    csv_var_names = list(items[0].keys())
    new_csv_var_names = compare_vars(csv_vars=csv_var_names, docx_vars=var_names, data=items)
    if not new_csv_var_names:
        return
    old_names = list(items[0].keys())
    renamed_items = []
    for item in items:
        new_item = {new_csv_var_names[i]: item[old_names[i]] for i in range(len(old_names))}
        renamed_items.append(new_item)
    items = renamed_items
    print(items)
    for n, item in enumerate(items):
        replace_placeholders(tree=tree,
                             xml_path=xml_path,
                             out_path=f"{out_path}/{file_name_pattern}".format(**item).replace(" ", "_"),
                             variables=item)