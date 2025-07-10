import streamlit as st
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Pt
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO

# Display Logo (uncomment and set path if needed)
# st.image("logo.png", width=200)

# Weekday and Holiday Setup
weekday_map = {
    "星期一": 0, "星期二": 1, "星期三": 2, "星期四": 3,
    "星期五": 4, "星期六": 5, "星期日": 6
}
weekday_chinese = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']

public_holidays = {
    "1 January 2025", "29 January 2025", "30 January 2025", "31 January 2025",
    "4 April 2025", "18 April 2025", "19 April 2025", "21 April 2025",
    "1 May 2025", "5 May 2025", "31 May 2025", "1 July 2025",
    "1 October 2025", "7 October 2025", "29 October 2025",
    "25 December 2025", "26 December 2025"
}
holiday_dates = set(datetime.strptime(d, "%d %B %Y").date() for d in public_holidays)

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
    "英文拼音", "高效寫字", "聆聽訓練", "說話訓練",
    "思維閱讀", "創意理解", "作文教學"
]

# Optional items configuration
optional_items_map = {
    "試堂日報讀贈券：即日報讀可獲舊生推薦現金券": -100,
    "試堂日報讀贈券：即日報讀可扣減試堂費": -200,
    "現金到校繳付24堂學費，送現金券": -50,
    "現金到校繳付36堂學費，送現金券": -50,
    "現金到校繳付48堂學費，送現金券": -100,
    "現金到校繳付72堂學費，送現金券": -100,
}

# Function Definitions
# ... (generate_schedule, calculate_week_range, calculate_main_course_fee, calculate_value_added_fee, calculate_optional_items, fill_template_doc)

# Streamlit UI
st.title(":calendar: 課程收據單生成器")

# User Inputs
student_name = st.text_input("學生姓名")
branch_name = st.selectbox("分校名稱", [
    "九龍灣(淘大)分校", "藍田(麗港城)分校", "青衣(青怡)分校",
    "九龍站(港景峯)分校", "鑽石山(萬迪廣場)分校"
])
invoice_number = st.text_input("單號")
total_lessons = st.selectbox("堂數", [4, 8, 10, 12, 24, 30, 36, 48, 72])

st.subheader("上課日及時間")
day_time_pairs = {}
for day in weekday_map:
    if st.checkbox(day):
        day_time_pairs[day] = st.selectbox(f"{day} 上課時間", lesson_time_options, key=day)

subjects = st.multiselect("主科", subject_options)
value_added_courses = st.multiselect("增值課程", value_added_options)
start_date = st.date_input("開始日期")

# Use the defined map for optional selections
optional_selections = st.multiselect("其他選項", list(optional_items_map.keys()))

# Generate Receipt
if st.button("生成收據單"):
    # Validate
    if not all([student_name, branch_name, invoice_number, subjects, day_time_pairs]):
        st.error("請填妥所有必填欄位。")
    else:
        # Compute schedules and fees
        lesson_dates, skipped_holidays = generate_schedule(
            total_lessons, list(day_time_pairs.keys()), start_date
        )
        week_range = calculate_week_range(
            total_lessons, len(day_time_pairs), lesson_dates
        )
        # Fees
        main_fee, main_material = calculate_main_course_fee(len(day_time_pairs), total_lessons)
        value_fee = calculate_value_added_fee(total_lessons)
        # assume no separate materials for value-added or adjust as needed
        value_material = 0
        opt_fee, opt_details = calculate_optional_items(optional_selections)
        total_amount = main_fee + main_material + value_fee + value_material + opt_fee

        # Fill and download document
        doc_file = fill_template_doc(
            student_name, branch_name, invoice_number,
            main_fee, main_material,
            value_fee, value_material,
            opt_details,
            start_date, lesson_dates, week_range,
            day_time_pairs, skipped_holidays,
            template_path
        )
        st.success("收據單已生成！")
        st.download_button(
            "下載 Word 文件", data=doc_file,
            file_name="課程收據單.docx"
        )

        # Clipboard Text
        end_date = start_date + timedelta(weeks=week_range) - timedelta(days=1)
        lines = [
            f"分校：{branch_name}",
            f"單號：{invoice_number}",
            f"學生姓名：{student_name}",
            f"堂數：{total_lessons}",
            f"學費金額：${total_amount}",
            f"主科：{' / '.join(subjects)}",
            f"增值課程：{' / '.join(value_added_courses)}",
            f"📆 上課期數範圍：{start_date.strftime('%d/%m/%Y')} 至 {end_date.strftime('%d/%m/%Y')}"
        ]
        lines += ["", "上課時間："] + [f"{day} {time}" for day, time in day_time_pairs.items()]
        lines += ["", "📅 上課日期安排："] + [
            f"{i}. {d.strftime('%d/%m/%Y')} ({weekday_chinese[d.weekday()]})"
            for i, d in enumerate(lesson_dates, 1)
        ]
        if skipped_holidays:
            lines += ["", "❌ 公眾假期 (休息):"] + [
                f"- {d.strftime('%d/%m/%Y')} ({weekday_chinese[d.weekday()]})" for d in skipped_holidays
            ]
        bill_text = '\n'.join(lines)
        st.subheader("📋 複製以下文字：")
        st.code(bill_text, language="text")



