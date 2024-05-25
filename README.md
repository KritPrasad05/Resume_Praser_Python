**Resume Parser**

**Project Overview**

The Resume Parser is a powerful tool designed to extract and organize crucial information from resumes in various formats. This project aims to automate the process of scanning resumes, identifying key details such as email addresses and phone numbers, and generating a structured Excel report for easy reference. It's an essential tool for recruiters, HR professionals, and anyone managing a large volume of resumes, helping them streamline their workflow and focus on strategic tasks.

**Features**

Multi-format Support: Capable of processing resumes in .doc, .docx, and .pdf formats.
Text Extraction: Utilizes advanced text extraction techniques to accurately capture resume content.
Contact Information Extraction: Automatically identifies and extracts email addresses and phone numbers.
Batch Processing: Converts .doc files to .docx format in batch mode for uniformity and ease of processing.
Report Generation: Creates a comprehensive Excel report summarizing candidate details, including extracted text for reference.

**Requirements**

Python 3.x
Libraries: pandas, docx, PyPDF2, win32com.client

**Installation**
1. Clone the repository:

git clone [https://github.com/yourusername/resume-parser.git](https://github.com/KritPrasad05/Resume_Praser_Python.git)
cd resume-parser

2. Install the required Python libraries:

pip install pandas python-docx PyPDF2 pywin32

**Usage**

Ensure your resume files (.doc, .docx, .pdf) are placed in a designated folder.

Run the script and provide the folder path containing your CV files when prompted:

python resume_parser.py

The script will:

Convert any .doc files to .docx format.
Process all .docx and .pdf files to extract text.
Identify and extract email addresses and phone numbers.
Generate an Excel report (cv_info.xlsx) containing the filename, email, contact number, and extracted text.

**Functions and Workflow**

convert_to_docx(doc_file): Converts a .doc file to .docx format using the win32com.client library.
batch_convert_to_docx(folder_path): Converts all .doc files in a specified folder to .docx format.
extract_text_from_pdf(pdf_path): Extracts text from a PDF file using the PyPDF2 library.
extract_text_from_word(docx_path): Extracts text from a .docx file using the docx library.
extract_info_from_text(text): Identifies email addresses and phone numbers from the extracted text using regular expressions.
process_cv_files(cv_folder): Processes all CV files in a specified folder, extracts information, and generates an Excel report.

Example

Enter the folder path containing CV files: path/to/your/cv/folder
CVs were scanned and the email, phone number, and overall text of the candidates were saved in the file 'cv_info.xlsx'.

**Contributing**

Contributions are welcome! If you have any suggestions, bug reports, or feature requests, feel free to create an issue or submit a pull request.

**Contact**

For any questions or feedback, please reach out to:- kritrp05@gmail.com
