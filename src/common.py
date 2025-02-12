from os import mkdir

import pandas as pd
import os
from pathlib import Path

dataset_dir = "dataset"
dataset_path = os.path.join(dataset_dir, "code.csv")
context_path = os.path.join(dataset_dir, "context.csv")

rag_dir = "rag_data"
rag_path = os.path.join(dataset_dir, "rag.csv")
rag_excel_path = os.path.join(dataset_dir, "rag.xlsx")

phrases = {
    "RAG_generate_old":
    """
Act as JavaScript and Excel VBA expert and code interpreter.Bellow you can find example of OnlyOffice API.
Please find Excel VBA equivalent for each method and answer in RAG text format providing only complete VBA and JS code.
Please keep in answer only a description of what the code does in the header and only code ( VBA and OnlyOffice JS).
This decription must be exist in two languages English and Russian Make comments in English for JavaScript and VBA code.
Following code:\n\n {prompt}
    """
    ,
    "RAG_replace_old":
    # maybe missing some important semantics here, for example function call arguments order
    """
Act as JavaScript and Excel VBA expert and code interpreter. Bellow you can find example of OnlyOffice API.
Remove specific details from code , replacing them symbols ###.
Please remove Description.
Please find Excel VBA equivalent for each method  and answer in RAG table with using JSON format.
Please keep in answer only a description of what the code does in the header and only code ( VBA and OnlyOffice JS) .
Following code:\n\n {prompt}
    """
    ,
    "RAG_filter_old":
    # LLM is not aware of knowledge it possesses !!!
    """
Act as JavaScript and Excel VBA expert and code interpreter. Bellow you can find RAG rules for converting VBA code.
Please keep here only important parts , which not represent in your network by default
Following code here:\n\n {prompt}
    """
    ,
    "Translate_preamble":
    """
Act as JavaScript and Excel VBA expert and code interpreter. 
Task is to translate code fragment from VBA using Excel API to JavaScript using OnlyOffice API.
    """
    ,
    "Translate_globals":
    """
These identifiers are defined outside the scope of code fragment and should not be modified:\n {prompt}.\n
    """
    ,
    "Translate_API":
    """
Use following samples to improve translation precision:\n {prompt}.\n 
    """
    ,
    "Translate_main_v1":
    """
Output should be only code and comments within it. 
If no translation possible place original code and the reason in comment. 
Code fragment written in Excel VBA to be translated:\n {prompt}.\n 
    """
    ,
    "Translate_main":
    """
Answer using code only. 
Code fragment written in Excel VBA to be translated:\n {prompt}.\n 
    """

}


# dataset generator
def get_dataset():
    code_df = pd.read_csv(dataset_path)
    context_df = pd.read_csv(context_path)
    context = {}

    for index, row in context_df.iterrows():
        fn = row["file_name"]
        ns = row["namespace"]
        val = row["names"]
        if fn not in context:
            context[fn] = {}
        context[fn][ns] = val

    for index, row in code_df.iterrows():
        fn = row["file_name"]
        ns = row["namespace"]
        val = row["code"]
        local_context = context[fn].get(ns, "")
        names: list = local_context.split(", ")
        for ns_i, names_str in context[fn].items():
            if ns_i == ns:
                continue
            ext_names = names_str.split(", ")
            for name in ext_names:
                names.append(ns_i + "." + name)
        present_names = []
        for name in names:
            if name in val:
                present_names.append(name)
        yield {
            "index": index,
            "code": val,
            "globals": present_names
        }


def extract_lang_code(text: str, lang_code: str) -> str:
    delim1 = f"```{lang_code}"
    delim2 = "```"
    if delim1 not in text:
        return ""
    val = text.split(delim1)[1]
    if delim2 not in val:
        return val
    val = val.split(delim2)[0]
    return val


def extract_js_code(text: str) -> str:
    return extract_lang_code(text, "javascript")


def extract_vba_code(text: str) -> str:
    return extract_lang_code(text, "vba")


class RunResult:

    def __init__(self, name: str):
        self.run_name = name
        self.temp_dir = f"data_temp\\{self.run_name}"

        Path(self.temp_dir).mkdir(parents=True, exist_ok=True)

    def log_question(self, index: int, text: str):
        file_path = os.path.join(self.temp_dir, f"{index}_question.txt")
        with open(file_path, "w") as text_file:
            text_file.write(text)

    def log_answer(self, index: int, text: str):
        file_path = os.path.join(self.temp_dir, f"{index}_answer.txt")
        with open(file_path, "w") as text_file:
            text_file.write(text)

    def save(self, results):
        dst_dir = "result"
        result_path = os.path.join(dst_dir, f"{self.run_name}.csv")
        result_excel_path = os.path.join(dst_dir, f"{self.run_name}.xlsx")
        df_result = pd.DataFrame(results, columns=["src_code", f"{self.run_name}", f"{self.run_name}_full_answer"])
        df_result.to_csv(result_path)
        df_result.to_excel(result_excel_path)
