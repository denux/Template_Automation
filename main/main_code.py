import pandas as pd
import json
import os
import pathlib
import glob
import sys
import re
import docx
from docxtpl import DocxTemplate
from docxcompose.composer import Composer



# some random comment, remote change
def getText(template_file_path):
    doc = docx.Document(template_file_path)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return ' '.join(fullText)

#this works
def exp1(asd):
    return asd

def generate_files(template_file_path, output_folder_path, input_dict):
    for count, local_dict in enumerate(input_dict):
        doc = DocxTemplate(template_file_path)
        doc.render(local_dict)
        doc.add_page_break()
        if count == 0:
            composer = Composer(doc)
        else:
            composer.append(doc)
    composer.save(output_folder_path + "output_doc.docx")


def get_config():
    with open(os.path.join(sys.path[0], "config.json"), "r") as f:
        dict_json = json.load(f)
        return dict_json


def main_fun():
    excel_folder = pathlib.Path(str(pathlib.Path().cwd().parent) + "/excel_data")
    excel_file_path = glob.glob(str(excel_folder) + "/*.*")[0]
    template_folder = pathlib.Path(str(pathlib.Path().cwd().parent) + "/template")
    template_file_path = glob.glob(str(template_folder) + "/*.*")[0]
    output_folder_path = str(pathlib.Path(str(pathlib.Path().cwd().parent) + "/template")) + "\\"
    extension_type = excel_file_path.split(".")[-1]
    if extension_type == "xlsx":
        df = pd.read_excel(excel_file_path)
    else:
        df = pd.read_csv(excel_file_path)

    temp_string = getText(template_file_path)
    unfiltered_columns = re.findall(r"\{\{[^}]+\}\}", temp_string)
    filtered_cols = [x.replace("{{", "").replace("}}", "").strip() for x in unfiltered_columns]
    if "filter" in df.columns:
        df = df[df["filter"] == "y"]

    df = df[filtered_cols]
    df.dropna(how="any", inplace=True)
    df = df.applymap(str)
    print(df)
    list_context = df.to_dict("records")
    generate_files(template_file_path, output_folder_path, list_context)


if __name__ == "__main__":
    main_fun()
