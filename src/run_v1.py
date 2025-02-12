from llm3 import get_response
import pandas as pd
import os
import common
from pprint import pprint
import datetime
from vbaL2 import VBA_L2
from antlr4 import *
from vbaLexer import vbaLexer
from vbaParser import vbaParser

os.chdir("..")

dst_dir = "result"
results = []

res = common.RunResult("run_v1_2")

rag_data = []
rag_index = {}  # keyword -> list of indexes from rag_data
rag_back_index = []  # list of keywords ordered same as rag_data

# returns list of indexes
def get_aug_data(src_code: str) -> list:
    res = []

    # parse
    lexer = vbaLexer(InputStream(src_code))
    stream = CommonTokenStream(lexer)
    parser = vbaParser(stream)
    query = VBA_L2()
    tree = parser.startRule()
    walker = ParseTreeWalker()
    walker.walk(query, tree)
    query.finalize()

    keywords = {}
    for name in query.names:
        keywords[name] = True
    while len(keywords) > 0:
        keyword = list(keywords)[0]
        if keyword in rag_index:
            i = rag_index[keyword][0]
            res.append(i)
            for bonus_keyword in rag_back_index[i]:
                if bonus_keyword in keywords:
                    del keywords[bonus_keyword]
        else:
            del keywords[keyword]
    return res


def load_aug_data():
    # ---------- load augmentation data ----------
    #  "file_name", "vba_code", "js_code", "ids", "comments"
    rag_df = pd.read_csv(common.rag_path)
    for index, row in rag_df.iterrows():
        i = len(rag_data)
        rag_rec = {
            "vba_code": row["vba_code"],
            "js_code": row["js_code"]
        }
        rag_data.append(rag_rec)
        keywords_str = str(row["ids"])
        keywords = keywords_str.split(", ")
        rag_back_index.append(keywords)
        for keyword in keywords:
            if keyword not in rag_index:
                rag_index[keyword] = []
            rag_index[keyword].append(i)


load_aug_data()
start_time = datetime.datetime.now()

for item in common.get_dataset():
    # if item["index"] == 0:
    #     continue
    prompt = common.phrases["Translate_preamble"]

    samples_i = get_aug_data(item["code"])

    if len(samples_i) > 0:
        samples_text = ""

        for i in range(len(samples_i)):
            sample = rag_data[samples_i[i]]
            samples_text += f"\nSample {i}:\n"
            samples_text += "VBA Excel code:\n"
            samples_text += sample["vba_code"]
            samples_text += "\n\nJavascript OnlyOffice code:\n"
            samples_text += sample["js_code"]
            samples_text += "\n"

        prompt += common.phrases["Translate_API"].format(prompt=samples_text)

    ignored_globals = item["globals"]
    if len(ignored_globals) > 0:
        prompt += common.phrases["Translate_globals"].format(prompt=", ".join(ignored_globals))

    prompt += common.phrases["Translate_main"].format(prompt=item["code"])

    res.log_question(index=item["index"], text=prompt)
    resp = get_response(prompt)
    res.log_answer(index=item["index"], text=resp)
    translated_code = common.extract_js_code(resp)

    res_record = [
        item["code"],
        translated_code,
        resp
    ]
    results.append(res_record)
    # if item["index"] > 2:
    #     break

res.save(results)

end_time = datetime.datetime.now()
duration = end_time - start_time
print("The duration is " + str(duration))





