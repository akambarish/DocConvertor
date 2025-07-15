#!/usr/bin/env python
# coding: utf-8

# In[ ]:





# In[1]:


# For categorizing files in their extension category.

import os
import shutil

# Source directory to search files
source_dir = r'C:\Users\ambar\Projects\PPT to PDFs\combined folder'  # <-- change this
# Output directory where categorized folders will be created
destination_dir = r'C:\Users\ambar\Projects\PPT to PDFs\seggregated-FIles'  # <-- change this

# Define file type mapping
file_types = {
    "Word_Files": [".doc", ".docx"],
    "PPT_Files": [".ppt", ".pptx"],
    "Excel_Files": [".xls", ".xlsx", ".xlsm"],
    "PDF_Files": [".pdf"]
}

# Create destination folders if not exist
for folder_name in file_types.keys():
    os.makedirs(os.path.join(destination_dir, folder_name), exist_ok=True)

# Walk through all subdirectories and files
for root, _, files in os.walk(source_dir):
    for file in files:
        file_ext = os.path.splitext(file)[1].lower()
        source_path = os.path.join(root, file)

        for folder_name, extensions in file_types.items():
            if file_ext in extensions:
                dest_path = os.path.join(destination_dir, folder_name, file)

                # Handle duplicate filenames
                base_name, ext = os.path.splitext(file)
                counter = 1
                while os.path.exists(dest_path):
                    dest_path = os.path.join(destination_dir, folder_name, f"{base_name}_{counter}{ext}")
                    counter += 1

                shutil.copy2(source_path, dest_path)
                break  # move to next file once matched

print("Files have been categorized successfully.")


# In[ ]:




