import os
import requests
import re
import time
import logging
from bs4 import BeautifulSoup
import urllib.parse

# Configuration
# Primary URL format
BASE_URL = "https://learn.microsoft.com/en-us/office/vba/api/excel."
# Alternative URL formats to try if the primary one fails
ALT_URL_FORMATS = [
    "https://learn.microsoft.com/en-us/office/vba/api/excel-{0}",  # Hyphenated format
    "https://learn.microsoft.com/en-us/office/vba/api/excel.{0}",  # Standard format
    "https://learn.microsoft.com/en-us/office/vba/excel/concepts/{0}",  # Concepts format
    "https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/ff838238(v=office.14)?redirectedfrom=MSDN#{0}"  # Legacy format
]
MD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'md')
LOG_FILE = 'ms_docs_vba_downloader.log'
CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ms_docs_updated_cache.txt')

# Track processed files and URLs
PROCESSED_FILES = set()
CHECKED_URLS = set()  # Track URLs that have already been checked

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

def save_processed_files_cache():
    """Save the processed files cache to a file"""
    with open(CACHE_FILE, 'w', encoding='utf-8') as f:
        for file in PROCESSED_FILES:
            f.write(f"{file}\n")
    logger.info(f"Saved {len(PROCESSED_FILES)} entries to cache: {CACHE_FILE}")

def save_checked_urls_cache():
    """Save the checked URLs cache to a file"""
    url_cache_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'checked_urls_cache.txt')
    with open(url_cache_file, 'w', encoding='utf-8') as f:
        for url in CHECKED_URLS:
            f.write(f"{url}\n")
    logger.info(f"Saved {len(CHECKED_URLS)} checked URLs to cache: {url_cache_file}")

def load_processed_files_cache():
    """Load the processed files cache from a file"""
    global PROCESSED_FILES
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, 'r', encoding='utf-8') as f:
            PROCESSED_FILES = set(line.strip() for line in f if line.strip())
        logger.info(f"Loaded {len(PROCESSED_FILES)} entries from cache: {CACHE_FILE}")
    else:
        PROCESSED_FILES = set()
        logger.info("No cache file found, starting with empty set")

def load_checked_urls_cache():
    """Load the checked URLs cache from a file"""
    global CHECKED_URLS
    url_cache_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'checked_urls_cache.txt')
    if os.path.exists(url_cache_file):
        with open(url_cache_file, 'r', encoding='utf-8') as f:
            CHECKED_URLS = set(line.strip() for line in f if line.strip())
        logger.info(f"Loaded {len(CHECKED_URLS)} checked URLs from cache: {url_cache_file}")
    else:
        CHECKED_URLS = set()
        logger.info("No URL cache file found, starting with empty set")

def get_items_to_process(directory):
    """Get all items (subfolders and md files) recursively in the given directory and all its descendants"""
    items = []
    
    # Walk through all directories and files recursively
    for root, dirs, files in os.walk(directory):
        # Process MD files in the current directory
        for file in files:
            if file.endswith('.md'):
                file_path = os.path.join(root, file)
                # Extract the base name without extension
                base_name = os.path.splitext(file)[0]
                # Use relative path from the base directory for display purposes
                rel_path = os.path.relpath(file_path, directory)
                items.append(('file', base_name, file_path, rel_path))
        
        # Process subfolders
        for dir_name in dirs:
            dir_path = os.path.join(root, dir_name)
            # Use relative path from the base directory for display purposes
            rel_path = os.path.relpath(dir_path, directory)
            items.append(('folder', dir_name, dir_path, rel_path))
    
    # Sort items by relative path for better organization
    items.sort(key=lambda x: x[3])
    
    print(f"Found {len(items)} items (files and folders) in directory tree")
    print(f"- {sum(1 for item in items if item[0] == 'file')} files")
    print(f"- {sum(1 for item in items if item[0] == 'folder')} folders")
    
    return items

def download_page(url):
    """Download a web page with retry logic"""
    global CHECKED_URLS

    time.sleep(2) 
    
    # Check if URL has already been checked
    if url in CHECKED_URLS:
        print("\n" + "="*80)
        print(f"SKIPPING ALREADY CHECKED URL: {url}")
        print("="*80)
        return None  # Return None for already checked URLs that failed
    
    # Display the URL prominently in the console
    print("\n" + "="*80)
    print(f"DOWNLOADING URL: {url}")
    print("="*80)
    
    # Add URL to checked URLs set
    CHECKED_URLS.add(url)
    
    # Set headers to simulate the latest Edge browser on Windows (which Microsoft definitely supports)
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36 Edg/121.0.0.0',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
        'Accept-Language': 'en-US,en;q=0.9',
        'Accept-Encoding': 'gzip, deflate, br',
        'Connection': 'keep-alive',
        'Upgrade-Insecure-Requests': '1',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'none',
        'Sec-Fetch-User': '?1',
        'Sec-Ch-Ua': '"Microsoft Edge";v="121", "Not A(Brand";v="24", "Chromium";v="121"',
        'Sec-Ch-Ua-Mobile': '?0',
        'Sec-Ch-Ua-Platform': '"Windows"',
        'Cache-Control': 'max-age=0',
        'DNT': '1'
    }
    
    max_retries = 3
    for attempt in range(max_retries):
        try:
            logger.info(f"Downloading: {url}")
            response = requests.get(url, headers=headers, timeout=30)
            if response.status_code == 200:
                print(f"Download successful (Status: {response.status_code})")
                
                # Show a snippet of the downloaded content
                html_content = response.text
                content_preview = html_content[:500] + "..." if len(html_content) > 500 else html_content
                print("\nDOWNLOADED CONTENT PREVIEW:")
                print("-"*80)
                print(content_preview)
                print("-"*80)
                
                return html_content
            else:
                error_msg = f"Failed to download {url}, status code: {response.status_code}"
                logger.warning(error_msg)
                print(f"Attempt {attempt+1}/{max_retries}: {error_msg}")
                if attempt < max_retries - 1:
                    print(f"Retrying in 2 seconds...")
                    time.sleep(2)  # Wait before retrying
        except Exception as e:
            error_msg = f"Error downloading {url}: {str(e)}"
            logger.error(error_msg)
            print(f"Attempt {attempt+1}/{max_retries}: {error_msg}")
            if attempt < max_retries - 1:
                print(f"Retrying in 2 seconds...")
                time.sleep(2)  # Wait before retrying
    
    print("Download failed after all retry attempts")
    return None

def extract_vba_code(html_content):
    """Extract VBA code examples from HTML content"""
    if not html_content:
        return None
    
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Look for code blocks with VBA content
    code_blocks = []
    
    # Method 1: Look for pre tags with VBA class or data-language attribute
    for pre in soup.find_all('pre'):
        if pre.get('data-language') == 'vba' or 'vba' in pre.get('class', []):
            code_blocks.append(pre.text)
        # Also check for code tags inside pre
        for code in pre.find_all('code'):
            if code.get('data-language') == 'vba' or 'vba' in code.get('class', []):
                code_blocks.append(code.text)
    
    # Method 2: Look for code tags with VBA class or data-language attribute
    for code in soup.find_all('code'):
        if code.get('data-language') == 'vba' or 'vba' in code.get('class', []):
            code_blocks.append(code.text)
    
    # Method 3: Look for div with code-example class that might contain VBA
    for div in soup.find_all('div', class_='code-example'):
        if 'vba' in div.get('data-language', '').lower() or 'vb' in div.get('data-language', '').lower():
            code_blocks.append(div.text)
    
    # If no code blocks found with specific VBA markers, try to find any code that looks like VBA
    if not code_blocks:
        # Look for any pre tags that might contain VBA code
        for pre in soup.find_all('pre'):
            text = pre.text.strip()
            # Check if it looks like VBA (contains common VBA keywords)
            vba_keywords = ['Sub ', 'Function ', 'Dim ', 'Set ', 'End Sub', 'End Function', 'Private ', 'Public ']
            if any(keyword in text for keyword in vba_keywords):
                code_blocks.append(text)
    
    # Show the extracted VBA code
    if code_blocks:
        print("\nEXTRACTED VBA CODE:")
        print("-"*80)
        for i, code in enumerate(code_blocks):
            print(f"\nCODE BLOCK #{i+1}:")
            print("-"*40)
            print(code[:300] + "..." if len(code) > 300 else code)
            print("-"*40)
    else:
        print("\nNo VBA code found in the downloaded content")
    
    return code_blocks if code_blocks else None

def extract_page_content(html_content):
    """Extract relevant content from the page and format as Markdown"""
    if not html_content:
        return None
    
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Check if the page contains the "browser not supported" message
    browser_not_supported = False
    if soup.find(text="This browser is no longer supported."):
        browser_not_supported = True
        print("\nWARNING: The page contains 'This browser is no longer supported' message.")
        print("Attempting to extract content from Microsoft's documentation archive.")
    
    # Extract title
    title = ""
    title_elem = soup.find('h1')
    if title_elem:
        title = title_elem.text.strip()
    else:
        # If no h1 found, try to construct a title from the URL or page content
        meta_title = soup.find('meta', property='og:title')
        if meta_title:
            title = meta_title.get('content', '')
        else:
            # Last resort: use the last part of the URL path
            title = "Excel VBA API Reference"
    
    # Extract description
    description = ""
    if browser_not_supported:
        # If browser not supported, provide a standard description
        description = "This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues."
    else:
        desc_elem = soup.find('p')
        if desc_elem:
            description = desc_elem.text.strip()
    
    # Extract VBA code examples
    vba_code = extract_vba_code(html_content)
    
    # Create markdown content
    md_content = f"# {title}\n\n"
    
    if description:
        md_content += f"## Description\n{description}\n\n"
    
    # Add syntax section if available
    syntax_section = soup.find('h2', string=re.compile('Syntax', re.I))
    if syntax_section and syntax_section.find_next('pre'):
        syntax_code = syntax_section.find_next('pre').text.strip()
        md_content += f"## Syntax\n```vba\n{syntax_code}\n```\n\n"
    
    # Add parameters section if available
    params_section = soup.find('h2', string=re.compile('Parameters|Arguments', re.I))
    if params_section:
        md_content += "## Parameters\n"
        params_table = params_section.find_next('table')
        if params_table:
            rows = params_table.find_all('tr')
            for row in rows[1:]:  # Skip header row
                cells = row.find_all(['td', 'th'])
                if len(cells) >= 2:
                    param_name = cells[0].text.strip()
                    param_desc = cells[1].text.strip()
                    md_content += f"- **{param_name}**: {param_desc}\n"
        md_content += "\n"
    
    # Add return value section if available
    return_section = soup.find('h2', string=re.compile('Return Value', re.I))
    if return_section:
        return_text = ""
        next_elem = return_section.find_next(['p', 'div'])
        if next_elem:
            return_text = next_elem.text.strip()
        md_content += f"## Return Value\n{return_text}\n\n"
    
    # Add remarks section if available
    remarks_section = soup.find('h2', string=re.compile('Remarks', re.I))
    if remarks_section:
        remarks_text = ""
        next_elem = remarks_section.find_next(['p', 'div'])
        if next_elem:
            remarks_text = next_elem.text.strip()
        md_content += f"## Remarks\n{remarks_text}\n\n"
    
    # Add example section
    md_content += "## Example\n"
    if vba_code:
        for code in vba_code:
            md_content += f"```vba\n{code}\n```\n\n"
    else:
        md_content += "No VBA example available.\n"
    
    # Show the extracted Markdown content
    print("\nEXTRACTED MARKDOWN CONTENT:")
    print("-"*80)
    print(md_content)
    print("-"*80)
    
    return md_content

def save_vba_code_to_bas_file(md_file_path, vba_code):
    """Save VBA code to a separate .exmp.bas file"""
    if not vba_code:
        return False
    
    # Create the .exmp.bas filename based on the MD file path
    base_name = os.path.splitext(md_file_path)[0]
    bas_file_path = f"{base_name}.exmp.bas"
    
    try:
        # Combine all code blocks into one file
        combined_code = "\n\n' ===== Next Example =====\n\n".join(vba_code)
        
        with open(bas_file_path, 'w', encoding='utf-8') as f:
            f.write(combined_code)
        
        # Display the saved VBA code in the console
        print("\n" + "="*80)
        print(f"SAVED VBA CODE FOR: {os.path.basename(bas_file_path)}")
        print("="*80)
        for i, code in enumerate(vba_code):
            print(f"\nEXAMPLE #{i+1}:")
            print("-"*40)
            print(code)
            print("-"*40)
        print("="*80 + "\n")
        
        logger.info(f"Saved VBA code to {bas_file_path}")
        return True
    except Exception as e:
        logger.error(f"Error saving VBA code to {bas_file_path}: {str(e)}")
        return False

def create_or_update_md_file(md_file_path, md_content):
    """Create or update the MD file with the extracted content"""
    try:
        with open(md_file_path, 'w', encoding='utf-8') as f:
            f.write(md_content)
        
        # Display the saved content in the console
        print("\n" + "="*80)
        print(f"SAVED MARKDOWN CONTENT FOR: {os.path.basename(md_file_path)}")
        print("="*80)
        print(md_content)
        print("="*80 + "\n")
        
        logger.info(f"Updated {md_file_path} with content from Microsoft docs")
        return True
    except Exception as e:
        logger.error(f"Error updating {md_file_path}: {str(e)}")
        return False

def process_item(item_type, item_name, item_path, rel_path):
    """Process an item (file or folder) to download and extract content"""
    # Determine the MD file path based on item type
    time.sleep(2) 
    if item_type == 'file':
        # For files, use the provided path directly
        md_file_path = item_path
    else:  # folder
        # For folders, find or create an MD file in the folder
        folder_path = item_path
        
        # Look for existing MD files
        md_files = []
        if os.path.exists(folder_path):
            md_files = [f for f in os.listdir(folder_path) if f.endswith('.md')]
        
        if not md_files:
            # Create a new MD file if none exists
            md_file = f"{item_name}.md"
            md_file_path = os.path.join(folder_path, md_file)
        else:
            # Update the first MD file found
            md_file_path = os.path.join(folder_path, md_files[0])
    
    # Check if the file has already been processed
    if md_file_path in PROCESSED_FILES:
        logger.info(f"Skipping already processed file: {md_file_path}")
        return True
    
    # Construct the URL - exclude the last part delimited by a dot
    url_part = item_name.replace(' ', '')  # Remove spaces
    
    print("\n" + "="*80)
    print(f"PROCESSING ITEM: {item_name}")
    print("="*80)
    
    # Remove the last part after the last dot (e.g., "ModelTableColumn.DataType.Property" -> "ModelTableColumn.DataType")
    if '.' in url_part:
        original_part = url_part
        url_part = '.'.join(url_part.split('.')[:-1])
        print(f"Modified URL part: {original_part} -> {url_part}")
    
    # Try the primary URL format first
    direct_url = BASE_URL + urllib.parse.quote(url_part)
    print(f"Attempting to download from primary URL: {direct_url}")
    logger.info(f"Trying URL: {direct_url}")
    html_content = download_page(direct_url)
    # Save the checked URLs cache after each URL check
    save_checked_urls_cache()
    
    # If primary URL fails, try with just the object name (e.g., "Range")
    if not html_content and '.' in url_part:
        object_name = url_part.split('.')[0]
        object_url = BASE_URL + urllib.parse.quote(object_name)
        print(f"Primary URL failed, trying object name URL: {object_url}")
        logger.info(f"Direct URL failed, trying object URL: {object_url}")
        html_content = download_page(object_url)
        # Save the checked URLs cache after each URL check
        save_checked_urls_cache()
    
    # If still no content, try all alternative URL formats
    if not html_content:
        print("\nPrimary URL formats failed. Trying alternative URL formats...")
        
        # Try different ways to format the URL part
        url_variations = [
            url_part.lower(),  # lowercase
            url_part.lower().replace('.', '-'),  # lowercase with hyphens instead of dots
            url_part.lower().replace('.', '/'),  # lowercase with slashes instead of dots
            url_part.lower().split('.')[-1] if '.' in url_part else url_part.lower(),  # just the last part
            url_part.lower().split('.')[0] if '.' in url_part else url_part.lower()   # just the first part
        ]
        
        # Try each URL format with each variation
        for url_format in ALT_URL_FORMATS:
            for variation in url_variations:
                if not html_content:  # Only try if we haven't found content yet
                    try:
                        alt_url = url_format.format(variation)
                        print(f"Trying alternative URL: {alt_url}")
                        logger.info(f"Trying alternative URL: {alt_url}")
                        html_content = download_page(alt_url)
                        # Save the checked URLs cache after each URL check
                        save_checked_urls_cache()
                        
                        if html_content and "This browser is no longer supported" not in html_content:
                            print(f"Alternative URL format worked: {alt_url}")
                            break
                    except Exception as e:
                        print(f"Error with URL format {url_format} and variation {variation}: {str(e)}")
                        continue
    
    if not html_content:
        logger.warning(f"Could not download content for {item_name}")
        return False
    
    # Check if the page contains the "browser not supported" message
    if html_content and "This browser is no longer supported" in html_content:
        print("\nWARNING: The page returned 'This browser is no longer supported' message.")
        
        # Try to bypass the browser check by modifying the request
        print("Attempting to bypass browser compatibility check...")
        
        # Try with a different User-Agent that specifically identifies as Edge
        edge_headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36 Edg/121.0.0.0',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.9',
            'Accept-Encoding': 'gzip, deflate, br',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
            'Sec-Fetch-Dest': 'document',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-Site': 'none',
            'Sec-Fetch-User': '?1',
            'Sec-Ch-Ua': '"Microsoft Edge";v="121", "Not A(Brand";v="24", "Chromium";v="121"',
            'Sec-Ch-Ua-Mobile': '?0',
            'Sec-Ch-Ua-Platform': '"Windows"',
            'DNT': '1'
        }
        
        # Try the Microsoft Docs mobile version URL which might have different browser checks
        mobile_url = direct_url.replace("learn.microsoft.com", "docs.microsoft.com/en-us")
        print(f"Trying mobile version URL: {mobile_url}")
        
        try:
            response = requests.get(mobile_url, headers=edge_headers, timeout=30)
            if response.status_code == 200:
                mobile_content = response.text
                if "This browser is no longer supported" not in mobile_content:
                    print("Mobile version URL worked!")
                    html_content = mobile_content
        except Exception as e:
            print(f"Error trying mobile version: {str(e)}")
        
        # Try the archive.org version as a last resort
        if "This browser is no longer supported" in html_content:
            archive_url = f"https://web.archive.org/web/20220101/{direct_url}"
            print(f"Trying archive.org URL: {archive_url}")
            
            try:
                response = requests.get(archive_url, headers=edge_headers, timeout=30)
                if response.status_code == 200:
                    archive_content = response.text
                    if "This browser is no longer supported" not in archive_content:
                        print("Archive.org version worked!")
                        html_content = archive_content
            except Exception as e:
                print(f"Error trying archive.org version: {str(e)}")
        
        # If we still have the browser not supported message, try to extract content anyway
        if "This browser is no longer supported" in html_content:
            print("All bypass attempts failed. Attempting to extract content despite browser compatibility issues...")
    
    # Extract VBA code from the HTML content
    vba_code = extract_vba_code(html_content)
    
    # Extract content and format as Markdown
    md_content = extract_page_content(html_content)
    if not md_content:
        logger.warning(f"Could not extract content for {item_name}")
        return False
    
    # Create or update the MD file
    md_success = create_or_update_md_file(md_file_path, md_content)
    
    # Save VBA code to a separate .exmp.bas file if available
    if vba_code:
        bas_success = save_vba_code_to_bas_file(md_file_path, vba_code)
        logger.info(f"VBA code extraction {'succeeded' if bas_success else 'failed'}")
    else:
        logger.info(f"No VBA code found for {item_name}")
    
    # If MD file update was successful, add to the processed files cache and save
    if md_success:
        PROCESSED_FILES.add(md_file_path)
        save_processed_files_cache()
    
    return md_success

def main():
    """Main function to process all items (files and folders)"""
    logger.info("Starting content download from Microsoft documentation")
    
    # Load the processed files and checked URLs caches
    load_processed_files_cache()
    load_checked_urls_cache()
    
    # Get all items in the MD directory
    items = get_items_to_process(MD_DIR)
    logger.info(f"Found {len(items)} items to process")
    
    # Show how many items are already processed
    already_processed = sum(1 for item_type, item_name, item_path, rel_path in items
                           if (item_type == 'file' and item_path in PROCESSED_FILES) or
                              (item_type == 'folder' and any(os.path.join(item_path, f) in PROCESSED_FILES
                                                            for f in os.listdir(item_path) if f.endswith('.md'))))
    logger.info(f"{already_processed} items already processed according to cache")
    
    # Process each item with interactive control
    success_count = 0
    proceed_all = False
    
    for i, (item_type, item_name, item_path, rel_path) in enumerate(items):
        item_desc = f"{item_type} '{rel_path}'"
        logger.info(f"Item {i+1}/{len(items)}: {item_desc}")
        
        if not proceed_all:
            print("\nOptions:")
            print(f"1 - Process next item: {item_desc}")
            print("2 - Process all remaining items")
            print("3 - Skip this item")
            print("4 - Exit")
            
            choice = input("Your choice (1/2/3/4): ").strip()
            
            if choice == '2':
                proceed_all = True
            elif choice == '3':
                logger.info(f"Skipping item: {item_desc}")
                continue
            elif choice == '4':
                logger.info("Exiting as requested by user")
                break
            elif choice != '1':
                print("Invalid choice. Defaulting to option 1 (process next)")
        
        logger.info(f"Processing {item_desc}")
        if process_item(item_type, item_name, item_path, rel_path):
            success_count += 1
        
        # Add a small delay to avoid overwhelming the server
        time.sleep(1)
    
    # Save the checked URLs cache
    save_checked_urls_cache()
    
    logger.info(f"Completed processing. Successfully updated {success_count} items.")

if __name__ == "__main__":
    main()