from ollama import chat
from ollama import ChatResponse


def get_response(content: str) -> str:

    response: ChatResponse = chat(
        model='hf.co/bartowski/DeepSeek-R1-Distill-Qwen-32B-GGUF:Q8_0',
        keep_alive="10m",
        messages=[
           {
              'role': 'user',
              'content': content,
           },
        ]
    )

    res: str = response['message']['content']
    res = res.replace("<think>","")
    res = res.replace("</think>", "")
    res = res.strip()
    return res

