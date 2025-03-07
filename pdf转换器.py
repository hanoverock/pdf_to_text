import fitz  # PyMuPDF
from docx import Document

def pdf_to_word(pdf_path, word_path):
    # Open the PDF file
    pdf_document = fitz.open(pdf_path)
    
    # Create a new Word document
    doc = Document()
    
    # Iterate through each page of the PDF
    for page_num in range(len(pdf_document)):
        # Extract text from the current page
        page = pdf_document.load_page(page_num)
        text = page.get_text()
        
        # Add the extracted text to the Word document
        doc.add_paragraph(text)
        
    # Save the Word document
    doc.save(word_path)
    
    # Close the PDF document
    pdf_document.close()

# Example usage
pdf_to_word(r'C:\Users\dawnl\Desktop\简历200517.pdf',r'C:\Users\dawnl\Desktop\Doc2.docx')
