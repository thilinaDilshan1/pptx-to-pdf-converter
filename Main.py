from pptx import Presentation
from reportlab.pdfgen import canvas

def pptx_to_pdf(pptx_path, pdf_path):
    prs = Presentation(pptx_path)
    pdf = canvas.Canvas(pdf_path)

    for slide_number, slide in enumerate(prs.slides):
        pdf.drawString(100, 800, f"Slide {slide_number + 1}")
        
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text
                pdf.drawString(100, 800 - (shape.top * 72), text)

        pdf.showPage()

    pdf.save()

pptx_to_pdf('E:/Campus/Y2/S2/CO 2225 - Software management techniques/Chapter6 - Project Human Resources Management.pptx', 'output.pdf')
