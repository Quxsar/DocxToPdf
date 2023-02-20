#In this code, we use "win32com.client" to control Word through "COM (Component Object Model)"
import win32com.client as win32
import os
#In the "convert_to_pdf" function, we open the file "file_path",
# change the extension to ".pdf", save the document in the new format and close it.
# Then we exit "Word"
def convert_to_pdf(file_path):
    word = win32.Dispatch('Word.Application')
    doc = word.Documents.Open(file_path)
    pdf_path = os.path.splitext(file_path)[0] + ".pdf"
    doc.SaveAs(pdf_path, FileFormat=17)
    doc.Close()
    word.Quit()

if __name__ == '__main__':
    #enter here the path to the file to be converted
    file_path = r"D:\maindirectory\test.docx"
    convert_to_pdf(file_path)