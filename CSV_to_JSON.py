#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import csv
import json
import os

def csv_to_json(csv_file_path, json_file_path):
    try:
        data = []
        with open(csv_file_path, mode='r', encoding='utf-8') as csv_file:
            reader = csv.DictReader(csv_file)
            for row in reader:
                data.append(row)

        with open(json_file_path, mode='w', encoding='utf-8') as json_file:
            json.dump(data, json_file, indent=4)

        print(f"Successfully converted '{csv_file_path}' to '{json_file_path}'.")

    except Exception as e:
        print(f"Error: {e}")

# Example usage
csv_file = 'example.csv'         # Path to your CSV file
json_file = 'output.json'        # Desired output JSON path
csv_to_json(csv_file, json_file)

