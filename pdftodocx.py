from pdf2docx import Converter


old_pdf="pdf_file.pdf"
new_doc= "new.docx"

obj = Converter(old_pdf)
obj.convert(new_doc)

obj.close()
