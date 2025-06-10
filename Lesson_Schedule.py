from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import os

def create_word_doc(student_name, branch_name, invoice_number, amount, total_lessons,
                    subjects, value_added_courses, lesson_times, start_date, lesson_dates):
    doc = Document()

    # Check for logo image existence and add it
    if os.path.exists("logo.png"):
        doc.add_picture("logo.png", width=Inches(2))
        doc.add_paragraph()

    def add_colored_text(paragraph, text, color_rgb, bold=False, size=16):
        run = paragraph.add_run(text)
        font = run.font
        font.size = Pt(size)
        font.color.rgb = RGBColor(*color_rgb)
        font.bold = bold

    # Title
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_colored_text(title, "Creat Learning\n創憶學坊", (0, 128, 0), True, 24)

    # Branch
    branch = doc.add_paragraph()
    branch.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_colored_text(branch, f"{branch_name} 分校", (0, 0, 255), False, 18)

    doc.add_paragraph()  # spacing

    # Student Info Section
    p = doc.add_paragraph()
    add_colored_text(p, "學生姓名：", (0, 0, 0), True)
    add_colored_text(p, f"{student_name}\n", (255, 0, 0))

    p = doc.add_paragraph()
    add_colored_text(p, "單號：", (0, 0, 0), True)
    add_colored_text(p, f"{invoice_number}\n", (255, 0, 0))

    p = doc.add_paragraph()
    add_colored_text(p, "金額：$", (0, 0, 0), True)
    add_colored_text(p, f"{amount}\n", (255, 0, 0))

    p = doc.add_paragraph()
    add_colored_text(p, "堂數：", (0, 0, 0), True)
    add_colored_text(p, f"{total_lessons}\n", (255, 0, 0))

    doc.add_paragraph()  # spacing

    # Course Info
    p = doc.add_paragraph()
    add_colored_text(p, "主科：", (0, 0, 0), True)
    add_colored_text(p, f"{subjects}\n", (128, 0, 128))

    p = doc.add_paragraph()
    add_colored_text(p, "增值課程：", (0, 0, 0), True)
    add_colored_text(p, f"{value_added_courses}\n", (128, 0, 128))

    p = doc.add_paragraph()
    add_colored_text(p, "上課時間：", (0, 0, 0), True)
    add_colored_text(p, f"{lesson_times}\n", (128, 0, 128))

    doc.add_paragraph()  # spacing

    # Start Date
    start_date_str = start_date.strftime('%d/%m/%Y')
    p = doc.add_paragraph()
    add_colored_text(p, "開始日期：", (0, 0, 0), True)
    add_colored_text(p, f"{start_date_str}\n", (255, 0, 0))

    doc.add_paragraph()  # spacing

    # Lesson Dates
    p = doc.add_paragraph()
    add_colored_text(p, "上課日期：\n", (0, 0, 0), True)

    for i, date in enumerate(lesson_dates, 1):
        date_str = date.strftime('%d/%m/%Y')
        date_para = doc.add_paragraph(f"{i}. {date_str}")
        date_para.paragraph_format.left_indent = Inches(0.3)

    # Save to BytesIO
    file_stream = BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream


