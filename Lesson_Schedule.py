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

# Function Definitions
...
# (Include generate_schedule, calculate_week_range, calculate_main_course_fee, calculate_value_added_fee, calculate_optional_items, fill_template_doc)

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
optional_selections = st.multiselect("其他選項", list(
    calculate_optional_items({}).__defaults__[0].keys()
))

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

def generate_schedule(total_lessons, frequency_days, start_date):
    freq_idxs = set(weekday_map[d] for d in frequency_days)
    lessons, skipped = [], []
    date = start_date
    while len(lessons) < total_lessons:
        if date.weekday() in freq_idxs:
            if date in holiday_dates:
                skipped.append(date)
            else:
                lessons.append(date)
        date += timedelta(days=1)
    return lessons, skipped


def calculate_week_range(total_lessons, freq_per_week, lesson_dates):
    if total_lessons in (10, 30):
        return total_lessons
    key = freq_per_week if freq_per_week < 3 else 3
    m = {
        1: {4:5,12:15,24:30},
        2: {8:5,24:15,48:30},
        3: {12:5,36:15,72:30}
    }
    base = m.get(key, {}).get(total_lessons, 5)
    holidays = sum(1 for d in lesson_dates if d in holiday_dates)
    return base + holidays


def calculate_main_course_fee(lessons_per_week, total_lessons):
    pricing = {
        (1, 4): (1280, 50), (1,12): (3456,100), (1,24): (6144,150),
        (2,8): (2400,100), (2,24):(5760,150),(2,48):(10080,300),
        (3,12):(3360,100),(3,36):(7560,250),(3,72):(14112,400),
        (None,10):(3500,100),(None,30):(9000,150)
    }
    return pricing.get((lessons_per_week, total_lessons)) or pricing.get((None,total_lessons),(0,0))


def calculate_value_added_fee(total_lessons):
    if total_lessons in (4,8): return 100 * total_lessons
    if total_lessons == 12:     return 75 * total_lessons
    if total_lessons == 24:     return 50 * total_lessons
    return 0


def calculate_optional_items(selected):
    fee, details = 0, []
    for opt in selected:
        amt = None
        if opt in optional_items_map:
            amt = optional_items_map[opt]
        elif "（＋$" in opt:
            amt = int(opt.split("（＋$")[-1].replace("）",""))
        if amt is not None:
            fee += amt
            details.append((opt, amt))
    return fee, details


def fill_template_doc(student_name, branch_name, invoice_number, main_tuition, main_material,
                      value_tuition, value_material, optional_items, start_date,
                      lesson_dates, week_range, day_time_pairs, skipped_holidays,
                      template_path):
    doc = Document(template_path)
    # Replace header fields
    reps = {
        "單號:": f"單號: {invoice_number}",
        "學生姓名：": f"學生姓名：{student_name}",
        "分校": f"分校：{branch_name}"
    }
    for p in doc.paragraphs:
        for k,v in reps.items():
            if p.text.strip().startswith(k): p.text = v

    # Insert fee calculation section
    fee_idx = next((i for i,p in enumerate(doc.paragraphs) if p.text.strip().startswith("學費計算")), None)
    if fee_idx is not None:
        # Main and materials
        doc.paragraphs[fee_idx].add_run("")
        doc.insert_paragraph(fee_idx+1, f"主科：+${main_tuition}")
        doc.insert_paragraph(fee_idx+2, f"小組活動教材：+${main_material}")
        # Value-added
        doc.insert_paragraph(fee_idx+3, f"增值課程學費：+${value_tuition}")
        doc.insert_paragraph(fee_idx+4, f"增值課程教材：+${value_material}")
        # Other
        doc.insert_paragraph(fee_idx+5, "其他:")
        for opt,amt in optional_items:
            doc.insert_paragraph(fee_idx+6, f"{opt}：{'+' if amt>0 else ''}${amt}")
        # Total
        total = main_tuition+main_material+value_tuition+value_material+sum(a for _,a in optional_items)
        doc.insert_paragraph(fee_idx+7, f"總額：= ${total}")

    # ... rest of document population (schedule, etc.) ...
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# Streamlit UI remains: call calculate_main_course_fee to get 4 values and pass to fill function




