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

lesson_time_options = [
    "9:30-11:00", "10:00-11:30", "10:30-12:00", "11:00-12:30",
    "11:30-13:00", "12:00-13:30", "13:30-15:00", "14:00-15:30",
    "14:30-16:00", "15:00-16:30", "15:30-17:00", "16:00-17:30",
    "16:30-18:00", "17:00-18:30", "17:30-19:00"
]

subject_options = [
    "中文記憶閱讀", "英文拼音", "小一面試班", "小學銜接班", "小學精進班"
]

value_added_options = [
    "英文拼音", "高效寫字", "聆聽訓練", "說話訓練", "思維閱讀", "創意理解", "作文教學"
]

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
    holiday_count = sum(1 for d in lesson_dates if d in holiday_dates)
    week_range += holiday_count
    return week_range

def create_word_doc(student_name, branch_name, invoice_number, amount, total_lessons,
                    subjects, value_added_courses, start_date,
                    lesson_dates, week_range, day_time_pairs):
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

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_colored_text(title, "Creat Learning\n創憶學坊", (0, 128, 0), True, 24)

    branch = doc.add_paragraph()
    branch.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_colored_text(branch, f"{branch_name} 分校", (0, 0, 255), False, 18)
    doc.add_paragraph()

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

    p = doc.add_paragraph()
    add_colored_text(p, "主科：", (0, 0, 0), True)
    add_colored_text(p, f"{' / '.join(subjects)}\n", (128, 0, 128))

    p = doc.add_paragraph()
    add_colored_text(p, "增值課程：", (0, 0, 0), True)
    add_colored_text(p, f"{' / '.join(value_added_courses)}\n", (128, 0, 128))
    doc.add_paragraph()

    start_date_str = start_date.strftime('%d/%m/%Y')
    p = doc.add_paragraph()
    add_colored_text(p, "開始日期：", (0, 0, 0), True)
    add_colored_text(p, f"{start_date_str}\n", (255, 0, 0))

    end_date = start_date + timedelta(weeks=week_range) - timedelta(days=1)
    p = doc.add_paragraph()
    add_colored_text(p, "上課期數範圍：", (0, 0, 0), True)
    add_colored_text(p, f"{start_date.strftime('%d/%m/%Y')} 至 {end_date.strftime('%d/%m/%Y')}\n", (0, 0, 0))
    doc.add_paragraph()

    p = doc.add_paragraph()
    add_colored_text(p, "上課日期：\n", (0, 0, 0), True)
    for i, date in enumerate(lesson_dates, 1):
        date_str = date.strftime('%d/%m/%Y')
        weekday_str = weekday_chinese[date.weekday()]
        time_str = day_time_pairs.get(weekday_str, "")
        date_para = doc.add_paragraph(f"{i}. {date_str} ({weekday_str}) {time_str}")
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

day_time_pairs = {}
st.subheader("上課日及時間")
for day in weekday_map.keys():
    if st.checkbox(f"{day}"):
        time = st.selectbox(f"選擇 {day} 上課時間", lesson_time_options, key=day)
        day_time_pairs[day] = time

subjects = st.multiselect("主科", subject_options)
value_added_courses = st.multiselect("增值課程", value_added_options)

start_date = st.date_input("開始日期", format="YYYY-MM-DD")

if st.button("生成收據單"):
    if all([student_name, branch_name, invoice_number, amount, subjects, day_time_pairs]):
        selected_days = list(day_time_pairs.keys())
        lesson_dates = generate_schedule(total_lessons, selected_days, start_date)
        week_range = calculate_week_range(total_lessons, len(selected_days), lesson_dates)
        doc_file = create_word_doc(student_name, branch_name, invoice_number, amount,
                                   total_lessons, subjects, value_added_courses,
                                   start_date, lesson_dates, week_range, day_time_pairs)

        st.success("收據單已生成！")
        st.download_button("📥 下載 Word 文件", data=doc_file, file_name="課程收據單.docx")
    else:
        st.error("請填妥所有必填欄位。")

