#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os
import shutil
from collections import defaultdict

def count_files_by_extension(root_folder):
    counts = defaultdict(int)
    file_paths = defaultdict(list)

    for foldername, subfolders, filenames in os.walk(root_folder):
        for filename in filenames:
            ext = os.path.splitext(filename)[1].lower()
            if ext:  # ignore files with no extension
                full_path = os.path.join(foldername, filename)
                counts[ext] += 1
                file_paths[ext].append(full_path)

    return counts, file_paths

def move_files_by_extension(file_paths, destination_root):
    for ext, paths in file_paths.items():
        dest_folder = os.path.join(destination_root, ext.strip('.'))
        os.makedirs(dest_folder, exist_ok=True)

        for path in paths:
            try:
                shutil.move(path, dest_folder)
            except Exception as e:
                print(f"Error moving {path}: {e}")

def count_files_in_sorted_folders(destination_root):
    sorted_counts = {}
    for foldername in os.listdir(destination_root):
        full_path = os.path.join(destination_root, foldername)
        if os.path.isdir(full_path):
            count = len([f for f in os.listdir(full_path) if os.path.isfile(os.path.join(full_path, f))])
            sorted_counts[foldername] = count
    return sorted_counts

# --------- Main Execution ---------
if __name__ == "__main__":
    source_folder = r"C:\Users\ambar\Projects\PPT to PDFs\combined folder"         # Replace with your source folder
    sorted_folder = r"C:\Users\ambar\Projects\PPT to PDFs\seggregated-FIles"           # Replace with your desired output folder

    print("Counting original files...")
    original_counts, file_paths = count_files_by_extension(source_folder)
    total_original = sum(original_counts.values())

    print("\nMoving files into sorted folders...")
    move_files_by_extension(file_paths, sorted_folder)

    print("\nCounting files in sorted folders...")
    sorted_counts = count_files_in_sorted_folders(sorted_folder)
    total_sorted = sum(sorted_counts.values())

    # Print summary
    print("\nOriginal File Counts by Extension:")
    for ext, count in sorted(original_counts.items()):
        print(f"{ext}: {count}")
    
    print("\nSorted File Counts by Folder:")
    for folder, count in sorted(sorted_counts.items()):
        print(f"{folder}: {count}")

    print(f"\n Total files originally: {total_original}")
    print(f" Total files after sorting: {total_sorted}")

    if total_original == total_sorted:
        print(" File count matches. All files sorted successfully.")
    else:
        print(" Mismatch in counts. Please verify.")


# In[ ]:




