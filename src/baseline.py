from operator import index

from llm3 import get_response
import pandas as pd
import os
import common
from pprint import pprint
import datetime

os.chdir("..")
dst_dir = "result"
results = []
res = common.RunResult("baseline_2")

start_time = datetime.datetime.now()

for item in common.get_dataset():
    # pprint(item)
    prompt = common.phrases["Translate_preamble"]
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





