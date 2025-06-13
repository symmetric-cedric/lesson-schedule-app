import streamlit as st
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Pt
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO

# Display Logo
st.image("logo.png", width=400)

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

template_path = "Testing.docx"

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

def generate_schedule(total_lessons, frequency_days, start_date):
    frequency_indices = sorted([weekday_map[day] for day in frequency_days])
    lessons = []
    skipped_holidays = []
    current_date = start_date

    while len(lessons) < total_lessons:
        for weekday in frequency_indices:
            days_ahead = (weekday - current_date.weekday() + 7) % 7
            lesson_date = current_date + timedelta(days=days_ahead)
            if lesson_date >= start_date:
                if lesson_date in holiday_dates:
                    skipped_holidays.append(lesson_date)
                else:
                    lessons.append(lesson_date)
                    if len(lessons) == total_lessons:
                        break
        current_date += timedelta(days=7)
    return lessons, skipped_holidays

def calculate_week_range(total_lessons, frequency_per_week, lesson_dates):
    if total_lessons == 10:
        return 10
    if total_lessons == 30:
        return 30
    key_freq = frequency_per_week if frequency_per_week < 3 else 3
    week_range_map = {
        1: {4: 5, 12: 15, 24: 30},
        2: {8: 5, 24: 15, 48: 30},
        3: {12: 5, 36: 15, 72: 30}
    }
    week_range = week_range_map.get(key_freq, {}).get(total_lessons, 5)
    holiday_count = sum(1 for d in lesson_dates if d in holiday_dates)
    return week_range + holiday_count

# Streamlit UI
st.title(":calendar: 課程收據單生成器")

student_name = st.text_input("學生姓名")
branch_name = st.selectbox("分校名稱", [
    "九龍灣(淘大)分校", "藍田(麗港城)分校", "青衣(青怡)分校", "九龍站(港景峯)分校", "鑽石山(萬迪廣場)分校"
])
invoice_number = st.text_input("單號")
amount = st.text_input("金額")
total_lessons = st.selectbox("堂數", [4, 8, 10, 12, 24, 30, 36, 48, 72])

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
        lesson_dates, skipped_holidays = generate_schedule(total_lessons, selected_days, start_date)
        week_range = calculate_week_range(total_lessons, len(selected_days), lesson_dates)
        end_date = start_date + timedelta(weeks=week_range) - timedelta(days=1)

        # Build text content for clipboard
        bill_text_lines = [
            f"分校：{branch_name}",
            f"單號：{invoice_number}",
            f"學生姓名：{student_name}",
            f"堂數：{total_lessons}",
            f"學費金額：${amount}",
            f"主科：{' / '.join(subjects)}",
            f"增值課程：{' / '.join(value_added_courses)}",
            f"📆 上課期數範圍：{start_date.strftime('%d/%m/%Y')} 至 {end_date.strftime('%d/%m/%Y')}",
            "",
            "📅 上課日期安排："
        ]
        for i, date in enumerate(lesson_dates, 1):
            weekday_str = weekday_chinese[date.weekday()]
            time_str = day_time_pairs.get(weekday_str, "")
            bill_text_lines.append(f"{i}. {date.strftime('%d/%m/%Y')} ({weekday_str}) {time_str}")

        if skipped_holidays:
            bill_text_lines.append("
❌ 公眾假期 (休息):")
            for d in skipped_holidays:
                bill_text_lines.append(f"- {d.strftime('%d/%m/%Y')} ({weekday_chinese[d.weekday()]})")
        else:
            bill_text_lines.append("
✅ 無需休息的公眾假期。")

        bill_text_lines.append("
📌 所有課程必須於限期內完成，逾期作廢。")
        bill_text = '
'.join(bill_text_lines)

        st.subheader("📋 複製以下文字：")
        st.code(bill_text, language="text")

        # Generate Word document
        def fill_template_doc(student_name, branch_name, invoice_number, amount, total_lessons,
                              subjects, value_added_courses, start_date,
                              lesson_dates, week_range, day_time_pairs, skipped_holidays):
            doc = Document(template_path)

            start_date_str = start_date.strftime('%d/%m/%Y')
            end_date = start_date + timedelta(weeks=week_range) - timedelta(days=1)
            date_range_str = f"{start_date_str} 至 {end_date.strftime('%d/%m/%Y')}"

            replacements = {
                "單號:": f"單號: {invoice_number}",
                "學生姓名：": f"學生姓名：{student_name}",
                "堂數：": f"堂數：{total_lessons}",
                "金額：": f"金額：${amount}",
                "主科": f"主科：{' / '.join(subjects)}",
                "增值課程": f"增值課程：{' / '.join(value_added_courses)}",
                "上課期數範圍": f"上課期數範圍：{date_range_str}",
                "分校": branch_name
            }

            for para in doc.paragraphs:
                for key, new_text in replacements.items():
                    if para.text.strip().startswith(key):
                        para.text = new_text

            insert_index = None
            for i, para in enumerate(doc.paragraphs):
                if "上課日期" in para.text:
                    insert_index = i + 1
                    break

            if insert_index is not None:
                doc.paragraphs.insert(insert_index, doc.add_paragraph(""))
                insert_index += 1

                table = doc.add_table(rows=1, cols=3)
                hdr_cells = table.rows[0].cells
                hdr_cells[0].text = "堂數"
                hdr_cells[1].text = "日期"
                hdr_cells[2].text = "時間"

                for cell in hdr_cells:
                    for paragraph in cell.paragraphs:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

                for i, date in enumerate(lesson_dates, 1):
                    row_cells = table.add_row().cells
                    row_cells[0].text = str(i)
                    row_cells[1].text = f"{date.strftime('%d/%m/%Y')} ({weekday_chinese[date.weekday()]})"
                    row_cells[2].text = day_time_pairs.get(weekday_chinese[date.weekday()], "")
                    for cell in row_cells:
                        for paragraph in cell.paragraphs:
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                doc.paragraphs[insert_index - 1]._element.addnext(table._element)

            for para in doc.paragraphs:
                if para.text.strip().startswith("公眾假期 (休息):"):
                    if skipped_holidays:
                        skipped_lines = ["公眾假期 (休息):"] + [
                            f"- {d.strftime('%d/%m/%Y')} ({weekday_chinese[d.weekday()]})" for d in skipped_holidays
                        ]
                        para.text = '
'.join(skipped_lines)
                    else:
                        para.text = ""
                    break

            file_stream = BytesIO()
            doc.save(file_stream)
            file_stream.seek(0)
            return file_stream

        doc_file = fill_template_doc(student_name, branch_name, invoice_number, amount,
                                     total_lessons, subjects, value_added_courses,
                                     start_date, lesson_dates, week_range, day_time_pairs, skipped_holidays)

        st.download_button("📅 下載 Word 文件", data=doc_file, file_name="課程收據單.docx")
        st.success("收據單已生成！")
    else:
        st.error("請填妥所有必填欄位。")


