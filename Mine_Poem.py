# extract_hafez_from_docx.py
from docx import Document
import re

DOCX_FILE = "divan_hafez.docx"  # مسیر فایل ورد شما
OUTPUT_FILE = "divan_hafez_list.py"

def clean_line(line):
    """
    پاکسازی خط: حذف شماره‌ها یا کاراکترهای غیرضروری
    """
    line = line.strip()
    # حذف شماره غزل (مثلاً "۱" یا "12")
    line = re.sub(r"^\d+\s*", "", line)
    return line

def extract_poems(docx_path):
    """
    استخراج غزل‌ها از فایل Word
    هر غزل به‌عنوان یک رشته در لیست قرار می‌گیرد
    """
    doc = Document(docx_path)
    poems = []
    current_poem = ""

    for para in doc.paragraphs:
        text = clean_line(para.text)
        if text == "":
            # خط خالی به معنی پایان غزل است
            if current_poem:
                poems.append(current_poem.strip())
                current_poem = ""
        else:
            current_poem += text + " "

    # اضافه کردن آخرین غزل در صورت وجود
    if current_poem:
        poems.append(current_poem.strip())

    return poems

def save_poems_list(poems, output_file):
    """
    ذخیره لیست غزل‌ها در قالب Python
    """
    with open(output_file, "w", encoding="utf-8") as f:
        f.write("HAFEZ_POEMS = [\n")
        for poem in poems:
            poem_escaped = poem.replace('"', '\\"')
            f.write(f'    "{poem_escaped}",\n')
        f.write("]\n")
    print(f"{len(poems)} غزل استخراج شد و در فایل {output_file} ذخیره شد.")

if __name__ == "__main__":
    poems = extract_poems(DOCX_FILE)
    save_poems_list(poems, OUTPUT_FILE)
