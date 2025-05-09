#!/usr/bin/env python3
"""
MD Folder Filter

This script reads a list of folder names from the excel_api file,
scans the MD directory for folders that start with these names followed by a dot,
and copies the matching folders to a new md_api directory.
"""

import os
import shutil
import sys
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('md_folder_filter.log'),
        logging.StreamHandler()
    ]
)

def read_excel_api(file_path):
    """
    Read the excel_api file and return a list of folder names.
    
    Args:
        file_path (str): Path to the excel_api file
        
    Returns:
        list: List of folder names
    """
    try:
        with open(file_path, 'r') as file:
            # Read lines and strip whitespace
            partial_names = [line.strip() for line in file if line.strip()]
        
        if not partial_names:
            logging.warning(f"The file {file_path} is empty. No folders will be filtered.")
        else:
            logging.info(f"Read {len(partial_names)} folder names from {file_path}")
            
        return partial_names
    
    except FileNotFoundError:
        logging.error(f"File not found: {file_path}")
        sys.exit(1)
    except Exception as e:
        logging.error(f"Error reading {file_path}: {str(e)}")
        sys.exit(1)

def get_folders_in_md(md_path):
    """
    Get all folders in the MD directory.
    
    Args:
        md_path (str): Path to the MD directory
        
    Returns:
        list: List of folder paths in the MD directory
    """
    try:
        if not os.path.isdir(md_path):
            logging.error(f"MD directory not found: {md_path}")
            sys.exit(1)
            
        # Get all folders in the MD directory
        folders = []
        for root, dirs, files in os.walk(md_path):
            for dir_name in dirs:
                # Get the full path of the directory
                full_path = os.path.join(root, dir_name)
                # Get the relative path from md_path
                rel_path = os.path.relpath(full_path, md_path)
                folders.append(rel_path)
                
        logging.info(f"Found {len(folders)} folders in {md_path}")
        return folders
    
    except Exception as e:
        logging.error(f"Error scanning MD directory: {str(e)}")
        sys.exit(1)

def filter_folders(all_folders, partial_names):
    """
    Filter folders based on partial names.
    
    Args:
        all_folders (list): List of all folder paths
        partial_names (list): List of partial folder names to match
        
    Returns:
        list: List of matching folder paths
    """
    matching_folders = []
    
    for folder in all_folders:
        folder_name = os.path.basename(folder)
        for partial_name in partial_names:
            # Check if folder name starts with partial_name followed by a dot
            if folder_name.startswith(f"{partial_name}."):
                matching_folders.append(folder)
                logging.info(f"Matched folder: {folder} (starts with '{partial_name}.')")
                break
    
    logging.info(f"Found {len(matching_folders)} matching folders out of {len(all_folders)}")
    return matching_folders

def create_md_api_directory(base_path):
    """
    Create the md_api directory.
    
    Args:
        base_path (str): Base path where md_api will be created
        
    Returns:
        str: Path to the created md_api directory
    """
    md_api_path = os.path.join(base_path, "md_api")
    
    try:
        # Check if directory already exists
        if os.path.exists(md_api_path):
            logging.warning(f"Directory already exists: {md_api_path}")
            # Ask user if they want to overwrite
            response = input(f"Directory {md_api_path} already exists. Overwrite? (y/n): ")
            if response.lower() != 'y':
                logging.info("Operation cancelled by user")
                sys.exit(0)
            # Remove existing directory
            shutil.rmtree(md_api_path)
            logging.info(f"Removed existing directory: {md_api_path}")
        
        # Create the directory
        os.makedirs(md_api_path)
        logging.info(f"Created directory: {md_api_path}")
        
        return md_api_path
    
    except Exception as e:
        logging.error(f"Error creating md_api directory: {str(e)}")
        sys.exit(1)

def copy_folders(matching_folders, md_path, md_api_path):
    """
    Copy matching folders from MD to md_api.
    
    Args:
        matching_folders (list): List of matching folder paths
        md_path (str): Path to the MD directory
        md_api_path (str): Path to the md_api directory
        
    Returns:
        int: Number of folders copied
    """
    copied_count = 0
    
    for folder in matching_folders:
        try:
            source_path = os.path.join(md_path, folder)
            dest_path = os.path.join(md_api_path, folder)
            
            # Create parent directories if they don't exist
            os.makedirs(os.path.dirname(dest_path), exist_ok=True)
            
            # Copy the directory and its contents
            shutil.copytree(source_path, dest_path)
            
            logging.info(f"Copied: {folder}")
            copied_count += 1
            
        except Exception as e:
            logging.error(f"Error copying {folder}: {str(e)}")
    
    return copied_count

def main():
    """
    Main function to orchestrate the process.
    """
    # Define paths
    base_path = os.getcwd()
    excel_api_path = os.path.join(base_path, "md_folder_filter_excel_api")
    md_path = os.path.join(base_path, "md")
    
    logging.info("Starting MD Folder Filter")
    logging.info(f"Base path: {base_path}")
    logging.info(f"excel_api path: {excel_api_path}")
    logging.info(f"MD path: {md_path}")
    
    # Read excel_api file
    partial_names = read_excel_api(excel_api_path)
    
    # Get all folders in MD directory
    all_folders = get_folders_in_md(md_path)
    
    # Filter folders based on partial names
    matching_folders = filter_folders(all_folders, partial_names)
    
    if not matching_folders:
        logging.warning("No matching folders found. Nothing to copy.")
        return
    
    # Create md_api directory
    md_api_path = create_md_api_directory(base_path)
    
    # Copy matching folders
    copied_count = copy_folders(matching_folders, md_path, md_api_path)
    
    logging.info(f"Completed. Copied {copied_count} folders to {md_api_path}")

if __name__ == "__main__":
    main()
