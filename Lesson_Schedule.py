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


def fill_template_doc(student_name, branch_name, invoice_number, amount, total_lessons,
                      subjects, value_added_courses, start_date,
                      lesson_dates, week_range, day_time_pairs, skipped_holidays,
                      optional_items, template_path):
    doc = Document(template_path)
    start_str = start_date.strftime('%d/%m/%Y')
    end = start_date + timedelta(weeks=week_range) - timedelta(days=1)
    range_str = f"{start_str} 至 {end.strftime('%d/%m/%Y')}"

    reps = {
        "單號:": f"單號: {invoice_number}",
        "學生姓名：": f"學生姓名：{student_name}",
        "堂數：": f"堂數：{total_lessons}",
        "學費金額：": f"學費金額：${amount}",
        "主科": f"主科：{' / '.join(subjects)}",
        "增值課程": f"增值課程：{' / '.join(value_added_courses)}",
        "上課期數範圍": f"上課期數範圍：{range_str}",
        "分校": f"分校：{branch_name}"
    }
    for p in doc.paragraphs:
        for k,v in reps.items():
            if p.text.strip().startswith(k): p.text=v

    # Insert schedule table
    idx = next((i for i,p in enumerate(doc.paragraphs) if "上課時間：" in p.text), None)
    if idx is not None:
        tbl = doc.add_table(rows=1, cols=3)
        hdr = tbl.rows[0].cells
        hdr[0].text, hdr[1].text, hdr[2].text = "堂數","日期","時間"
        for i,d in enumerate(lesson_dates,1):
            row = tbl.add_row().cells
            row[0].text = str(i)
            wd = weekday_chinese[d.weekday()]
            row[1].text = f"{d.strftime('%d/%m/%Y')} ({wd})"
            row[2].text = day_time_pairs.get(wd,"")
            for c in row:
                for par in c.paragraphs: par.alignment=WD_ALIGN_PARAGRAPH.CENTER
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        doc.paragraphs[idx]._element.addnext(tbl._element)

    if optional_items:
        doc.add_paragraph("\n其他項目:")
        for opt,amt in optional_items:
            doc.add_paragraph(f"{opt}：{'+' if amt>0 else ''}${amt}")

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# Streamlit UI
st.title(":calendar: 課程收據單生成器")

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

# Optional promotions & add-ons
optional_selections = st.multiselect("其他選項", list(optional_items_map.keys()))

if st.button("生成收據單"):
    if not all([student_name, branch_name, invoice_number, subjects, day_time_pairs]):
        st.error("請填妥所有必填欄位。")
    else:
        lesson_dates, skipped = generate_schedule(total_lessons, list(day_time_pairs.keys()), start_date)
        week_range = calculate_week_range(total_lessons, len(day_time_pairs), lesson_dates)
        main_fee, _ = calculate_main_course_fee(len(day_time_pairs), total_lessons)
        value_fee = calculate_value_added_fee(total_lessons)
        opt_fee, opt_details = calculate_optional_items(optional_selections)
        total_amount = main_fee + value_fee + opt_fee

        doc_file = fill_template_doc(
            student_name, branch_name, invoice_number, total_amount,
            total_lessons, subjects, value_added_courses, start_date,
            lesson_dates, week_range, day_time_pairs, skipped, opt_details, template_path
        )
        st.success("收據單已生成！")
        st.download_button("下載 Word 文件", data=doc_file, file_name="課程收據單.docx")

        end_date = start_date + timedelta(weeks=week_range) - timedelta(days=1)
        lines = [
            f"分校：{branch_name}", f"單號：{invoice_number}",
            f"學生姓名：{student_name}", f"堂數：{total_lessons}",
            f"學費金額：${total_amount}",
            f"主科：{' / '.join(subjects)}", f"增值課程：{' / '.join(value_added_courses)}",
            f"📆 上課期數範圍：{start_date.strftime('%d/%m/%Y')} 至 {end_date.strftime('%d/%m/%Y')}"
        ]
        lines += ["", "上課時間："] + [f"{d} {t}" for d,t in day_time_pairs.items()]
        lines += ["", "📅 上課日期安排："] + [
            f"{i}. {d.strftime('%d/%m/%Y')} ({weekday_chinese[d.weekday()]})"
            for i,d in enumerate(lesson_dates,1)
        ]
        if skipped:
            lines += ["", "❌ 公眾假期 (休息):"] + [
                f"- {d.strftime('%d/%m/%Y')} ({weekday_chinese[d.weekday()]})" for d in skipped
            ]
        bill_text = '\n'.join(lines)
        st.subheader("📋 複製以下文字：")
        st.code(bill_text, language="text")



