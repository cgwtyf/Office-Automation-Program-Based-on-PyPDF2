from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.pagesizes import letter
# from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
# from reportlab.platypus import Paragraph
from docx2pdf import convert
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# 注册字体文件
def register_font(font_path):
    pdfmetrics.registerFont(TTFont('SimSun', font_path))

# 添加水印到PDF文件
def add_watermark_to_pdf(input_pdf, output_pdf, watermark_text, rotation_angle, opacity):
    pdf_reader = PdfReader(input_pdf)
    pdf_writer = PdfWriter()

    watermark = canvas.Canvas("watermark.pdf", pagesize=letter)

    # 设置旋转角度
    watermark.rotate(rotation_angle)

    # 设置透明度
    # watermark.setFillAlpha(opacity)

    # 设置字体和字体大小
    watermark.setFont("SimSun", 70)

    # 设置字体颜色
    watermark.setFillColor(colors.gray)

    watermark.setFillGray(0.8)

    watermark.setFillAlpha(0.5)

    # 设置水印文本位置
    watermark.drawString(12*cm, 3*cm, watermark_text)

    watermark.save()

    watermark_pdf = PdfReader("watermark.pdf")
    watermark_page = watermark_pdf.pages[0]

    for page_num in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[page_num]

        # 创建一个新的页面，并将水印文件合并到新页面上
        new_page = page
        new_page.merge_page(watermark_page)

        pdf_writer.add_page(new_page)

    with open(output_pdf, "wb") as output:
        pdf_writer.write(output)

# 将Word文档转换为PDF并添加水印
def convert_word_to_pdf_with_watermark(input_docx, output_pdf, watermark_text, rotation_angle, opacity):
    # 将Word文档转换为PDF
    convert(input_docx, output_pdf)

    # 添加水印到PDF
    add_watermark_to_pdf(output_pdf, output_pdf, watermark_text, rotation_angle, opacity)

# 示例用法
input_docx = "动态规划算法-成国伟-教案设计.docx"
output_pdf = "output.pdf"
watermark_text = "好思享教育"
rotation_angle = 45  # 旋转角度，以度为单位
opacity = 0.1  # 透明度，范围从0到1

font_path = "simsun.ttf"  # 字体文件路径

register_font(font_path)

customChoice=int(input("选择您需要批量自动化的操作：\n1：word批量转pdf\n2：word批量转pdf并加水印\n"))
if customChoice==1:
    convert(input_docx, output_pdf)
elif customChoice==2:
    watermark_text = input("请输入你的水印文字：\n")
    convert_word_to_pdf_with_watermark(input_docx, output_pdf, watermark_text, rotation_angle, opacity)
