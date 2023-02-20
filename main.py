import win32com.client as win32
import os

def convert_to_pdf(file_path):
    word = win32.Dispatch('Word.Application')
    doc = word.Documents.Open(file_path)
    pdf_path = os.path.splitext(file_path)[0] + ".pdf"
    doc.SaveAs(pdf_path, FileFormat=17)
    doc.Close()
    word.Quit()

if __name__ == '__main__':
    # source file path example
    file_path = r"D:\maindirectory\test.docx"
    convert_to_pdf(file_path)