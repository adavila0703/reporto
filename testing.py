from fpdf import FPDF

pdf = FPDF()
pdf.add_page()
pdf.set_font('Arial', 'B', 16)
pdf.cell(40, 10, 'hi!')
pdf.cell(50, 10, 'yo!')
pdf.output('tuto1.pdf', 'F')