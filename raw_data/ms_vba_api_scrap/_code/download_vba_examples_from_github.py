import os
import requests
import re
import time
import logging
import argparse

# Configuration
GITHUB_API_URL = "https://api.github.com"
MD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'md')
os.makedirs(MD_DIR, exist_ok=True)
SUMMARY_LOG = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'search_summary.log')
CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'downloaded_files_cache.txt')

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('vba_downloader.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

HEADERS = {
    "Accept": "application/vnd.github.v3+json",
    "Authorization": "Token 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"
}

# Search tracking dictionary (flags only)
SEARCH_CACHE = {}

# Track downloaded files
DOWNLOADED_FILES = set()

def save_downloaded_files_cache():
    """Save the downloaded files cache to a file"""
    with open(CACHE_FILE, 'w', encoding='utf-8') as f:
        for file in DOWNLOADED_FILES:
            f.write(f"{file}\n")
    logger.info(f"Saved {len(DOWNLOADED_FILES)} entries to download cache: {CACHE_FILE}")

def load_downloaded_files_cache():
    """Load the downloaded files cache from a file"""
    global DOWNLOADED_FILES
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, 'r', encoding='utf-8') as f:
            DOWNLOADED_FILES = set(line.strip() for line in f if line.strip())
        logger.info(f"Loaded {len(DOWNLOADED_FILES)} entries from download cache: {CACHE_FILE}")
    else:
        DOWNLOADED_FILES = set()
        logger.info("No download cache found, starting with empty set")

def search_github_for_vba_examples(search_term):
    """Search GitHub for VBA examples with search tracking"""
    
    # Trim last segment after last dot
    processed_term = '.'.join(search_term.split('.')[:-1]) if '.' in search_term else search_term
    
    # Check if term has been searched before
    if processed_term in SEARCH_CACHE:
        logger.info(f"Already searched for: {processed_term}")
        return None  # Return None to indicate cached knowledge
    
    time.sleep(2)  # Rate limit protection
    search_query = f"{processed_term}++language%3Avba+&type=repositories&s=stars&o=desc"
    search_url = f"{GITHUB_API_URL}/search/code?q={search_query}"
    
    logger.info(f"Searching GitHub for: {processed_term} (original: {search_term})")
    logger.info(f"URL {search_url}")
    try:
        response = requests.get(search_url, headers=HEADERS)
        if response.status_code == 200:
            items = response.json()['items']
            logger.info(f"Found {len(items)} results")
            SEARCH_CACHE[processed_term] = True  # Mark as searched
            return items
        else:
            logger.error(f"API error {response.status_code} - waiting 30s")
            time.sleep(5)
            return None
    except Exception as e:
        logger.error(f"Search error: {str(e)}")
        return None

def download_vba_file(item, output_dir, base_name):
    """Download a VBA file"""
    try:
        download_url = item['html_url'].replace('github.com', 'raw.githubusercontent.com').replace('/blob/', '/')
        file_name = f"{base_name}_{os.path.basename(item['path'])}"
        local_path = os.path.join(MD_DIR, file_name)
        
        logger.info(f"Downloading: {file_name}")
        response = requests.get(download_url)
        
        if response.status_code == 200:
            with open(local_path, 'w', encoding='utf-8') as f:
                f.write(response.text)
            logger.info(f"Saved: {local_path}")
            # Update the cache file after each successful download
            DOWNLOADED_FILES.add(os.path.basename(local_path))
            save_downloaded_files_cache()
            return local_path
        else:
            logger.error(f"Download failed (Status: {response.status_code})")
            return None
    except Exception as e:
        logger.error(f"Download error: {str(e)}")
        return None

def log_search_summary(md_file, result_count):
    """Log search results to summary file"""
    with open(SUMMARY_LOG, 'a') as f:
        f.write(f"{md_file},{result_count}\n")

def has_vba_files(md_file):
    """Check if a markdown file already has associated VBA files (.cls or .bas)"""
    md_dir = os.path.dirname(md_file)
    base_name = os.path.splitext(os.path.basename(md_file))[0]
    
    # Check if there are any .cls or .bas files in the same directory with the same base name
    for file in os.listdir(md_dir):
        if file.endswith(('.cls', '.bas')) and base_name in file:
            logger.info(f"Found existing VBA file for {base_name}: {file}")
            return True
    
    return False

def process_md_files():
    """Process MD files with interactive control"""
    # No longer resetting DOWNLOADED_FILES as we want to keep the persistent cache
    
    with open(SUMMARY_LOG, 'w') as f:
        f.write("md_file,result_count\n")
    logger.info(f"Summary log: {os.path.abspath(SUMMARY_LOG)}")
    
    md_files = []
    for root, _, files in os.walk(MD_DIR):
        for file in files:
            if file.endswith('.md'):
                md_files.append(os.path.join(root, file))
    
    proceed_all = False
    for md_file in md_files:
        base_name = os.path.splitext(os.path.basename(md_file))[0]
        output_dir = os.path.join(os.path.dirname(md_file), base_name)
        os.makedirs(output_dir, exist_ok=True)
        
        # Check if the markdown file already has associated VBA files
        if has_vba_files(md_file):
            logger.info(f"Skipping {md_file} as it already has associated VBA files")
            log_search_summary(md_file, 0)
            continue
        
        if not proceed_all:
            print("\nOptions:")
            print("1 - Process next MD file")
            print("2 - Process all remaining")
            choice = input("Your choice (1/2): ").strip()
            if choice == '2':
                proceed_all = True
            elif choice != '1':
                return
        
        items = search_github_for_vba_examples(base_name)
        if items:
            for item in items[:3]:  # Max 3 downloads
                file_name = f"{base_name}_{os.path.basename(item['path'])}"
                if file_name not in DOWNLOADED_FILES:
                    download_vba_file(item, MD_DIR, base_name)
                    DOWNLOADED_FILES.add(file_name)
                    save_downloaded_files_cache()
            log_search_summary(md_file, len(items))
        else:
            log_search_summary(md_file, 0)

def main():
    # Load the downloaded files cache
    load_downloaded_files_cache()
    process_md_files()

if __name__ == "__main__":
    main()
