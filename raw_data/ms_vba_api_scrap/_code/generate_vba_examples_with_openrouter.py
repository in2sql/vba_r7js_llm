import os
import requests
import time

# Configuration
API_URL = "https://openrouter.ai/api/v1/chat/completions"
API_KEY = os.getenv("OPENROUTER_API_KEY")  # Set your API key as an environment variable
MODEL = "google/gemini-2.0-flash-001"
MD_ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'md')
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output_vba_examples')
os.makedirs(OUTPUT_DIR, exist_ok=True)

HEADERS = {
    "Authorization": f"Bearer {API_KEY}",
    "Content-Type": "application/json"
}

PROMPT_TEMPLATE = (
    "You are an expert in Excel VBA. Read the following documentation and generate a realistic, practical VBA example that demonstrates how to use the described methods or properties. "
    "Only output the VBA code, with brief inline comments. Do NOT include any other explanation.\n\n"
    "Documentation:\n{md_content}\n\nVBA Example:"
)

def generate_vba_example(md_content):
    data = {
        "model": MODEL,
        "messages": [
            {"role": "user", "content": PROMPT_TEMPLATE.format(md_content=md_content)}
        ],
        "max_tokens": 1000
    }
    response = requests.post(API_URL, headers=HEADERS, json=data)
    if response.status_code == 200:
        result = response.json()
        return result['choices'][0]['message']['content'].strip()
    else:
        print(f"API error: {response.status_code} {response.text}")
        return None

def classify_api_popularity(md_content):
    prompt = (
        "Based on the following Excel VBA API documentation, rate the popularity of this API among Excel VBA developers on a scale from 1 (least popular) to 10 (most popular). "
        "Only output a single integer from 1 to 10.\n\nDocumentation:\n" + md_content
    )
    data = {
        "model": MODEL,
        "messages": [
            {"role": "user", "content": prompt}
        ],
        "max_tokens": 4
    }
    response = requests.post(API_URL, headers=HEADERS, json=data)
    if response.status_code == 200:
        result = response.json()
        content = result['choices'][0]['message']['content'].strip()
        try:
            score = int(re.search(r'\d+', content).group())
            score = min(max(score, 1), 10)
        except Exception:
            score = 5  # Default
        return score
    else:
        print(f"API error (popularity): {response.status_code} {response.text}")
        return 5

def main():
    for root, dirs, files in os.walk(MD_ROOT):
        for fname in files:
            if not fname.endswith('.md'):
                continue
            md_path = os.path.join(root, fname)
            with open(md_path, 'r', encoding='utf-8') as f:
                md_content = f.read()
            rel_path = os.path.relpath(md_path, MD_ROOT)
            print(f"Generating VBA example for {rel_path}...")
            vba_example = generate_vba_example(md_content)
            api_score = classify_api_popularity(md_content)
            out_fname = f"{rel_path}.vba_{api_score}.md"
            out_path = os.path.join(OUTPUT_DIR, out_fname.replace(os.sep, '_'))
            if vba_example:
                with open(out_path, 'w', encoding='utf-8') as outf:
                    outf.write(f"# {rel_path} (popularity: {api_score}/10)\n\n")
                    outf.write(vba_example)
                    outf.write("\n\n---\n\n")
                print(f"Saved: {out_path}")
            time.sleep(5)  # Be polite to the API
    print(f"Done. All VBA examples written to {OUTPUT_DIR}")

if __name__ == "__main__":
    main()
