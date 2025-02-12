#  eats too much memory

from pprint import *
from transformers import pipeline

messages = [
    {"role": "user", "content": "Who are you?"},
]
pipe = pipeline("text-generation", model="deepseek-ai/DeepSeek-R1-Distill-Qwen-32B", device="cpu")
print("model loaded")
res = pipe(messages)
print("result:")
pprint(res)
