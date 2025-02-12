import os
from fileinput import filename

from antlr4.InputStream import InputStream

from antlr4 import *
from vbaLexer import vbaLexer
from vbaParser import vbaParser
from vbaL1 import VBA_L1
import pandas as pd
import common

from oletools.olevba import VBA_Parser

os.chdir("..")
src_dir = "source_data"
temp_dir = "data_temp"

code_data = []  # file_name, namespace, code
context_data = []  # file_name, namespace, list of identifiers


def fill_dataset(s: str, fn: str):
    lexer = vbaLexer(InputStream(s))
    stream = CommonTokenStream(lexer)
    parser = vbaParser(stream)
    query = VBA_L1()
    tree = parser.startRule()
    walker = ParseTreeWalker()
    walker.walk(query, tree)

    if len(query.global_code) > 0:
        global_code = "\n".join(query.global_code)
        query.subs.append(global_code)

    for sub in query.subs:
        rec = [fn, query.namespace, sub]
        code_data.append(rec)

    if len(query.ids)>0:
        context_rec = [
            fn,
            query.namespace,
            ", ".join(query.ids)
        ]
        context_data.append(context_rec)


for file_name in os.listdir(src_dir):
    print(f"src file {file_name}")
    file_path = os.path.join(src_dir, file_name)
    if not os.path.isfile(file_path):
        continue

    vba_parser = VBA_Parser(file_path)

    if not vba_parser.detect_vba_macros():
        print('No VBA Macros found')

    for (sub_filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
        dst_path = os.path.join(temp_dir, file_name + "." + vba_filename)
        with open(dst_path, "w") as text_file:
            text_file.write(vba_code)
        fill_dataset(vba_code, file_name)
    vba_parser.close()


df_code = pd.DataFrame(code_data, columns=["file_name", "namespace", "code"])
df_code.to_csv(common.dataset_path)

df_context = pd.DataFrame(context_data, columns=["file_name", "namespace", "names"])
df_context.to_csv(common.context_path)


