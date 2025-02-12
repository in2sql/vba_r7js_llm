## Transpiler for vba->js and Excel->OnlyOffice API 

featuring LLM, ANTRL and RAG

latest model used:
https://huggingface.co/QuantFactory/Qwen2.5-Coder-7B-Instruct-GGUF

### Usage:
1. backup everything from dataset if needed
2. Place xlsx files into "source_data" directory
3. run [prepare_src_data.py](src%2Fprepare_src_data.py)
4. [baseline.py](src%2Fbaseline.py) to get translation using llm only
5. [run_v1.py](src%2Frun_v1.py) to get translation using llm with RAG
6. check result folder

[data_temp](data_temp) contains intermediates (llm dialog, pure extracted data, etc.)

note: llama-cpp requires run as admin in windows
