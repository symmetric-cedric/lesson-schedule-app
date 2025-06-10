import streamlit as st
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Pt, RGBColor
from io import BytesIO

# Weekday and Holiday Setup
weekday_map = {
    "星期一": 0, "星期二": 1, "星期三": 2, "星期四": 3, "星期五": 4, "星期六": 5, "星期日": 6
}

weekday_chinese = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']

# Public Holidays in 2025 (Hong Kong)
public_holidays = {
    "1 January 2025", "29 January 2025", "30 January 2025", "31 January 2025",
    "4 April 2025", "18 April 2025", "19 April 2025", "21 April 2025",
    "1 May 2025", "5 May 2025", "31 May 2025", "1 July 2025", "1 October 2025",
    "7 October 2025", "29 October 2025", "25 December 2025", "26 December 2025"
}
holiday_dates = set(datetime.strptime(date_str, "%d %B %Y").date() for date_str in public_holidays)

# Schedule Function
def generate_schedule(total_lessons, frequency_days, start_date):
    frequency_indices = sorted([weekday_map[day] for day in frequency_days])
    lessons = []
    current_date = start_date

    while len(lessons) < total_lessons:
        for weekday in frequency_indices:
            days_ahead = (weekday - current_date.weekday() + 7) % 7
            lesson_date = current_date + timedelta(days=days_ahead)
            if lesson_date >= start_date and lesson_date not in holiday_dates:
                lessons.append(lesson_date)
                if len(lessons) == total_lessons:
                    break
        current_date += timedelta(days=7)
    return lessons

# Word Document Generator
def create_word_doc(student_name, branch_name, invoice_number, amount, total_lessons,
                    subjects, value_added_courses, lesson_times, start_date, lesson_dates):
    doc = Document()

    def add_colored_text(text, color_rgb, bold=False, size=16):
        run = doc.add_paragraph().add_run(text)
        run.font.size = Pt(size)
        run.font.color.rgb = RGBColor(*color_rgb)
        run.bold = bold

    # Header Info
    add_colored_text("Creat Learning\n創憶學坊", (0, 128, 0), True)
    add_colored_text(f"{branch_name} 分校", (0, 0, 255))

    doc.add_paragraph(f"學生姓名：").add_run(student_name).font.color.rgb = RGBColor(255, 0, 0)
    doc.add_paragraph(f"單號：").add_run(invoice_number).font.color.rgb = RGBColor(255, 0, 0)
    doc.add_paragraph(f"金額：$").add_run(amount).font.color.rgb = RGBColor(255, 0, 0)
    doc.add_paragraph(f"堂數：").add_run(str(total_lessons)).font.color.rgb = RGBColor(255, 0, 0)

    # Subjects
    p = doc.add_paragraph("主科：")
    run = p.add_run(subjects)
    run.font.color.rgb = RGBColor(128, 0, 128)

    # Value-Added Courses
    p = doc.add_paragraph("增值課程：")
    run = p.add_run(value_added_courses)
    run.font.color.rgb = RGBColor(128, 0, 128)

    # Lesson Times
    p = doc.add_paragraph("上課時間：")
    run = p.add_run(lesson_times)
    run.font.color.rgb = RGBColor(128, 0, 128)

    # Start Date (no parsing needed — just strftime)
    start_date_str = start_date.strftime('%d/%m/%Y')
    doc.add_paragraph("開始日期：").add_run(start_date_str).font.color.rgb = RGBColor(255, 0, 0)

    # Lesson Dates
    doc.add_paragraph("上課日期：")
    for i, date in enumerate(lesson_dates, 1):
        date_str = date.strftime('%d/%m/%Y')
        doc.add_paragraph(f"{i}. {date_str}")

    # Save to BytesIO
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
total_lessons = st.number_input("堂數", min_value=1, max_value=100, step=1)

subjects = st.text_area("主科（以 / 分隔）")
value_added_courses = st.text_area("增值課程（以 / 分隔）")
lesson_times = st.text_area("上課時間")

start_date = st.date_input("開始日期", format="YYYY-MM-DD")
frequency_options = list(weekday_map.keys())
selected_days = st.multiselect("上課日（可選多日）", frequency_options)

if st.button("生成收據單"):
    if all([student_name, branch_name, invoice_number, amount, subjects, lesson_times, selected_days]):
        lesson_dates = generate_schedule(total_lessons, selected_days, start_date)
        doc_file = create_word_doc(student_name, branch_name, invoice_number, amount,
                                   total_lessons, subjects, value_added_courses,
                                   lesson_times, start_date, lesson_dates)

        st.success("收據單已生成！")
        st.download_button("📥 下載 Word 文件", data=doc_file, file_name="課程收據單.docx")
    else:
        st.error("請填妥所有必填欄位。")



