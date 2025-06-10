import streamlit as st
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import os

# Weekday and Holiday Setup
weekday_map = {
    "星期一": 0, "星期二": 1, "星期三": 2, "星期四": 3, "星期五": 4, "星期六": 5, "星期日": 6
}
weekday_chinese = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']

public_holidays = {
    "1 January 2025", "29 January 2025", "30 January 2025", "31 January 2025",
    "4 April 2025", "18 April 2025", "19 April 2025", "21 April 2025",
    "1 May 2025", "5 May 2025", "31 May 2025", "1 July 2025", "1 October 2025",
    "7 October 2025", "29 October 2025", "25 December 2025", "26 December 2025"
}
holiday_dates = set(datetime.strptime(date_str, "%d %B %Y").date() for date_str in public_holidays)

# Functions

def generate_schedule(total_lessons, frequency_days, start_date):
    frequency_indices = sorted([weekday_map[day] for day in frequency_days])
    lessons = []
    current_date = start_date

    while len(lessons) < total_lessons:
        for weekday in frequency_indices:
            days_ahead = (weekday - current_date.weekday() + 7) % 7
            lesson_date = current_date + timedelta(days=days_ahead)
            if lesson_date >= start_date:
                lessons.append(lesson_date)
                if len(lessons) == total_lessons:
                    break
        current_date += timedelta(days=7)
    return lessons

def calculate_week_range(total_lessons, frequency_per_week, lesson_dates):
    key_freq = frequency_per_week if frequency_per_week < 3 else 3
    week_range_map = {
        1: {4: 5, 12: 15, 24: 30},
        2: {8: 5, 24: 15, 48: 30},
        3: {12: 5, 36: 15, 72: 30}
    }
    week_range = week_range_map.get(key_freq, {}).get(total_lessons, 5)

    # Add +1 week for each holiday overlap
    holiday_count = sum(1 for d in lesson_dates if d in holiday_dates)
    week_range += holiday_count
    return week_range

def create_word_doc(student_name, branch_name, invoice_number, amount, total_lessons,
                    subjects, value_added_courses, lesson_times, start_date,
                    lesson_dates, week_range):
    doc = Document()

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

    branch = doc.add_paragraph()
    branch.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_colored_text(branch, f"{branch_name} 分校", (0, 0, 255), False, 18)
    doc.add_paragraph()

    # Student Info
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
    doc.add_paragraph()

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
    doc.add_paragraph()

    # Start Date
    start_date_str = start_date.strftime('%d/%m/%Y')
    p = doc.add_paragraph()
    add_colored_text(p, "開始日期：", (0, 0, 0), True)
    add_colored_text(p, f"{start_date_str}\n", (255, 0, 0))

    # Week Range
    end_date = start_date + timedelta(weeks=week_range) - timedelta(days=1)
    p = doc.add_paragraph()
    add_colored_text(p, "上課期數範圍：", (0, 0, 0), True)
    add_colored_text(p, f"{start_date.strftime('%d/%m/%Y')} 至 {end_date.strftime('%d/%m/%Y')}\n", (0, 0, 0))

    doc.add_paragraph()

    # Lesson Dates
    p = doc.add_paragraph()
    add_colored_text(p, "上課日期：\n", (0, 0, 0), True)
    for i, date in enumerate(lesson_dates, 1):
        date_str = date.strftime('%d/%m/%Y')
        date_para = doc.add_paragraph(f"{i}. {date_str}")
        date_para.paragraph_format.left_indent = Inches(0.3)

    file_stream = BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

# Streamlit UI
st.title(":calendar: 課程收據單生成器")

student_name = st.text_input("學生姓名")
branch_name = st.selectbox("分校名稱", [
    "創憶學坊(淘大)", "創憶學坊(麗港城)", "創憶學坊(青衣)", "創憶學坊(港景峯)", "創憶學坊(鑽石山)"
])
invoice_number = st.text_input("單號")
amount = st.text_input("金額")
total_lessons = st.selectbox("堂數", [4, 8, 12, 24, 36, 48, 72])

subjects = st.text_area("主科（以 / 分隔）")
value_added_courses = st.text_area("增值課程（以 / 分隔）")
lesson_times = st.text_area("上課時間")

start_date = st.date_input("開始日期", format="YYYY-MM-DD")
frequency_options = list(weekday_map.keys())
selected_days = st.multiselect("上課日（可選多日）", frequency_options)

if st.button("生成收據單"):
    if all([student_name, branch_name, invoice_number, amount, subjects, lesson_times, selected_days]):
        lesson_dates = generate_schedule(total_lessons, selected_days, start_date)
        week_range = calculate_week_range(total_lessons, len(selected_days), lesson_dates)
        doc_file = create_word_doc(student_name, branch_name, invoice_number, amount,
                                   total_lessons, subjects, value_added_courses,
                                   lesson_times, start_date, lesson_dates, week_range)

        st.success("收據單已生成！")
        st.download_button("📥 下載 Word 文件", data=doc_file, file_name="課程收據單.docx")
    else:
        st.error("請填妥所有必填欄位。")



