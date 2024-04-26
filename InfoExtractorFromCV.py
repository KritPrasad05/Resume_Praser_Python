import os
import re
import pandas as pd
import docx
from PyPDF2 import PdfReader
import win32com.client


def convert_to_docx(doc_file):
    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(doc_file)
    new_file = os.path.splitext(doc_file)[0] + ".docx"
    doc.SaveAs(new_file, FileFormat=16)  # FileFormat=16 for .docx
    doc.Close()
    word.Quit()


def batch_convert_to_docx(folder_path):
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if file_name.lower().endswith(".doc"):
            convert_to_docx(file_path)


def extract_text_from_pdf(pdf_path):
    text = ""
    with open(pdf_path, "rb") as f:
        pdf_reader = PdfReader(f)
        for page in pdf_reader.pages:
            text += page.extract_text()
    return text


def extract_text_from_word(docx_path):
    text = ""
    doc = docx.Document(docx_path)
    for paragraph in doc.paragraphs:
        text += paragraph.text
    return text


def extract_info_from_text(text):
    email_set = set(re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b', text))
    phone = re.findall(r'(?:\b\+?91[\s-]?)?(?:(?:[6789]\d{2}[\s-]?)?\d{3}[\s-]?\d{4}\b)', text)
    return list(email_set), phone


def process_cv_files(cv_folder):
    existing_files = [file for file in os.listdir('.') if file.startswith('cv_info')]
    if existing_files:
        max_num = max([int(re.findall(r'\d+', file)[0]) for file in existing_files])
        new_file_name = f'cv_info{max_num + 1}.xlsx'
    else:
        new_file_name = 'cv_info1.xlsx'

    data = {'Filename': [], 'Email': [], 'Contact': [], 'Text': []}
    for file_name in os.listdir(cv_folder):
        if file_name.endswith('.pdf'):
            cv_path = os.path.join(cv_folder, file_name)
            text = extract_text_from_pdf(cv_path)
        elif file_name.endswith('.docx'):
            cv_path = os.path.join(cv_folder, file_name)
            text = extract_text_from_word(cv_path)
        else:
            continue
        email, contact = extract_info_from_text(text)
        data['Filename'].append(file_name)
        data['Email'].append(email)
        data['Contact'].append(contact)
        data['Text'].append(text)

    df = pd.DataFrame(data)
    df.to_excel(new_file_name, index=False)
    print(
        f"CVs were scanned and the email, phone number, and overall text of the candidates were saved in the file '{new_file_name}'.")


# Prompt the user to input the folder path containing CV files
cv_folder = input("Enter the folder path containing CV files: ")

# Convert .doc files to .docx
batch_convert_to_docx(cv_folder)

# Process CV files and extract information
process_cv_files(cv_folder)
