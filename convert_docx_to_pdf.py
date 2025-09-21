import os
import sys
import comtypes.client
from PyPDF2 import PdfReader, PdfWriter

def docx_to_pdf(docx_path, pdf_path, password):
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False
    try:
        doc = word.Documents.Open(docx_path, False, True, None, PasswordDocument=password)
        doc.SaveAs(pdf_path, FileFormat=17)  # 17 is the format for PDF
        doc.Close()
        print(f"Converted '{docx_path}' to '{pdf_path}' successfully.")
    except Exception as e:
        print(f"Failed to convert '{docx_path}' to PDF: {e}")
    finally:
        word.Quit()

def encrypt_pdf(input_pdf_path, output_pdf_path, password):
    try:
        reader = PdfReader(input_pdf_path)
        writer = PdfWriter()

        for page in reader.pages:
            writer.add_page(page)

        writer.encrypt(password)
        with open(output_pdf_path, 'wb') as f:
            writer.write(f)
        print(f"Encrypted PDF saved as '{output_pdf_path}'.")
    except Exception as e:
        print(f"Failed to encrypt PDF: {e}")

if __name__ == "__main__":
    docx_file = "PASSWORDS.docx"
    temp_pdf_file = "PASSWORDS_temp.pdf"
    final_pdf_file = "PASSWORDS.pdf"
    password = os.getenv("DOCX_PASSWORD")

    if not password:
        print("Error: Environment variable 'DOCX_PASSWORD' is not set.")
        sys.exit(1)

    if not os.path.exists(docx_file):
        print(f"File '{docx_file}' does not exist.")
        sys.exit(1)

    docx_to_pdf(os.path.abspath(docx_file), os.path.abspath(temp_pdf_file), password)
    encrypt_pdf(os.path.abspath(temp_pdf_file), os.path.abspath(final_pdf_file), password)
    os.remove(temp_pdf_file)
