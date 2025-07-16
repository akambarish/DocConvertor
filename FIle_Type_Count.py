#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os
from collections import defaultdict

def count_file_types(directory):
    """
    Counts the number of files of each type in a directory and its subdirectories.

    Args:
        directory (str): The path to the directory to scan.

    Returns:
        dict: A dictionary where keys are file extensions (e.g., '.docx')
              and values are the count of files with that extension.
              Returns None if the directory is not valid.
    """
    if not os.path.isdir(directory):
        print(f"Error: The specified path '{directory}' is not a valid directory.")
        return None

    # Use defaultdict to automatically handle the first time an extension is found
    file_counts = defaultdict(int)

    print(f"Scanning directory: {directory}...\n")

    # os.walk generates the file names in a directory tree
    for dirpath, _, filenames in os.walk(directory):
        for filename in filenames:
            # os.path.splitext splits the path into a (root, ext) pair
            # We are interested in the extension, which is the second element
            extension = os.path.splitext(filename)[1].lower()

            if extension:
                file_counts[extension] += 1
            else:
                # If there's no extension, we can count it as 'no_extension'
                file_counts['(no extension)'] += 1

    return file_counts

def print_summary(file_counts):
    """Prints a formatted summary of the file counts."""
    if not file_counts:
        print("No files found in the specified directory.")
        return

    print("--- File Type Count Summary ---")
    print(f"{'Extension':<20} {'Count'}")
    print("-" * 30)

    # Sort the items by count in descending order for better readability
    sorted_counts = sorted(file_counts.items(), key=lambda item: item[1], reverse=True)

    total_files = 0
    for extension, count in sorted_counts:
        print(f"{extension:<20} {count}")
        total_files += count

    print("-" * 30)
    print(f"{'Total Files Found:':<20} {total_files}")
    print("-----------------------------")


if __name__ == "__main__":
    # Get the target directory from the user
    # You can also hardcode a path here, for example:
    # target_directory = "C:/Users/YourUser/Documents"
    target_directory = input("Enter the full path of the folder you want to scan: ")

    # Run the counter
    counts = count_file_types(target_directory)

    # Print the results
    if counts is not None:
        print_summary(counts)


# In[ ]:




