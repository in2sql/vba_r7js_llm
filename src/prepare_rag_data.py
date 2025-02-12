import os
from fileinput import filename

from antlr4.InputStream import InputStream

from antlr4 import *
from vbaLexer import vbaLexer
from vbaParser import vbaParser
from vbaL2 import VBA_L2
import pandas as pd
import common
from tqdm import tqdm


os.chdir("..")
src_dir = "rag_data\\_txt"

rag_data = []  # file_name, vba_code, js_code, list of identifiers, comments


def fill_dataset(all_text: str, fn: str):
    vba_code = common.extract_vba_code(all_text)
    js_code = common.extract_js_code(all_text)

    lexer = vbaLexer(InputStream(vba_code))
    stream = CommonTokenStream(lexer)
    parser = vbaParser(stream)
    query = VBA_L2()
    tree = parser.startRule()
    walker = ParseTreeWalker()
    walker.walk(query, tree)
    query.finalize()

    rag_rec = [
        fn,
        vba_code,
        js_code,
        ", ".join(query.names),
        "\n".join(query.comments),
    ]
    rag_data.append(rag_rec)


file_names = os.listdir(src_dir)

for file_name in tqdm(file_names):
    file_path = os.path.join(src_dir, file_name)
    if not os.path.isfile(file_path):
        continue
    with open(file_path, 'r', encoding="utf-8") as file:
        file_text = file.read()
    fill_dataset(file_text, file_name)


df = pd.DataFrame(rag_data, columns=["file_name", "vba_code", "js_code", "ids", "comments"])
df.to_csv(common.rag_path, encoding="utf-8")
df.to_excel(common.rag_excel_path)
