# download tesseract exe here and install it ,addit to path variable enviroment i.e C:\Program Files\Tesseract-OCR
# https://github.com/UB-Mannheim/tesseract/wiki

#run file === (env) E:\projects\python\convert>python.exe e:/projects/python/convert.py


import pytesseract
from PIL import Image
import fitz  # PyMuPDF
from docx import Document
from io import BytesIO

# Path to the scanned PDF image
pdf_image_path = "C:\\Users\\user\\Desktop\\roadmap to address issues and recommendations.pdf"

# Path to save the Word document
output_docx_path =  "C:\\Users\\user\\Desktop\\output_word_document.docx"

# Extract text from the scanned PDF image using OCR
def extract_text_from_pdf_image(pdf_image_path):
    text = ''
    with fitz.open(pdf_image_path) as pdf:
        for page_num in range(len(pdf)):
            page = pdf.load_page(page_num)
            image_list = page.get_images(full=True)
            for img_index, img in enumerate(image_list):
                xref = img[0]
                base_image = pdf.extract_image(xref)
                image_bytes = base_image["image"]
                # Create a file-like object from the image bytes
                image_file = BytesIO(image_bytes)
                # Open the image using PIL
                image = Image.open(image_file)
                text += pytesseract.image_to_string(image)
    return text

# Create a Word document and write the extracted text into it
def create_word_document(text, output_docx_path):
    document = Document()
    document.add_paragraph(text)
    document.save(output_docx_path)

# Extract text from the scanned PDF image
extracted_text = extract_text_from_pdf_image(pdf_image_path)

# Create a Word document and write the extracted text into it
create_word_document(extracted_text, output_docx_path)

print(f'Successfully converted scanned PDF image to Word document: {output_docx_path}')
