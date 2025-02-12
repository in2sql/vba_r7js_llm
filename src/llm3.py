from llama_cpp import Llama
from pprint import pprint

llm = Llama.from_pretrained(
    repo_id="QuantFactory/Qwen2.5-Coder-7B-Instruct-GGUF",
    filename="Qwen2.5-Coder-7B-Instruct.Q8_0.gguf",
    n_gpu_layers=-1,
    n_ctx=32768
)


def get_response(content: str) -> str:
    # response = llm(
    #     content,
    #     max_tokens=None,
    #     echo=False
    # )
    #
    # pprint(response)
    # res: str = response['choices'][0]['text']
    response = llm.create_chat_completion(
        messages=[
            {
                "role": "user",
                "content": content
            }
        ],
        max_tokens=None
    )
    # pprint(response)
    res: str = response['choices'][0]['message']['content']
    res = res.strip()
    return res

# print(get_response("Who are you?"))
