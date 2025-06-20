import os
import re
import shutil
import subprocess
from docx import Document

# Define folder containing Word files (absolute path)
folder_path = os.path.join(os.path.dirname(__file__), "./../word_files")

# Stopwords and stemmer
mongolian_stopwords = ["бол", "энэ", "тэр", "байна", "юм", "гээд", "гэх", "үү", "байдаг", "өөр", "гэж", "нь"]

def simple_mongolian_stem(word):
    suffixes = ["ийн", "ын", "тэй", "гүй", "ууд", "уудын", "д", "т", "аар", "оор", "ыг", "ийг"]
    for suffix in suffixes:
        if word.endswith(suffix):
            return word[:-len(suffix)]
    return word

def clean_text(text):
    text = text.lower()
    text = re.sub(r'[^\w\s]', '', text)
    text = re.sub(r'\d+', '', text)
    text = ' '.join(text.split())
    tokens = text.split()
    tokens = [word for word in tokens if word not in mongolian_stopwords]
    tokens = [simple_mongolian_stem(word) for word in tokens]
    return ' '.join(tokens)

# Loop through all Word files
for filename in os.listdir(folder_path):
    if filename.endswith(".docx"):
        file_path = os.path.join(folder_path, filename)
        try:
            doc = Document(file_path)
            full_text = "\n".join([para.text for para in doc.paragraphs])
            cleaned = clean_text(full_text)
            print(f"\n🧼 Cleaned Text from {filename}:\n{cleaned}\n")
            # (Optional) Save cleaned version to new file
            with open(os.path.join(folder_path, f"cleaned_{filename.replace('.docx', '.txt')}"), "w", encoding="utf-8") as f:
                f.write(cleaned)
        except Exception as e:
            print(f"Failed to process {filename}: {e}")
