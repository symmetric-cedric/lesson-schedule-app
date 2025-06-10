import streamlit as st
from datetime import datetime, timedelta

# Weekday and Holiday Setup
weekday_map = {
    "星期一": 0,
    "星期二": 1,
    "星期三": 2,
    "星期四": 3,
    "星期五": 4,
    "星期六": 5,
    "星期日": 6
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

def calculate_weeks_range(total_lessons, lessons_per_week, lessons_dates):
    base_weeks = (total_lessons / lessons_per_week) * 5 / 4
    base_weeks = int(base_weeks) if base_weeks == int(base_weeks) else int(base_weeks) + 1
    
    # Count lessons overlapping holidays
    holiday_count = sum(1 for d in lessons_dates if d in holiday_dates)
    
    total_weeks = base_weeks + holiday_count
    
    return total_weeks

def generate_schedule(total_lessons, frequency_days, start_date, student_name, school_name):
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

    formatted = [
        f"{i+1}. {student_name} | {dt.year}年 {dt.month} 月 {dt.day} 日 ({weekday_chinese[dt.weekday()]}) | {school_name}分校"
        for i, dt in enumerate(lessons)
    ]
    return formatted

# Streamlit UI
st.title(":calendar: 課程日期安排")

student_name = st.text_input("Student Name")
school_name = st.selectbox("Select School Branch", [
    "創憶學坊(淘大)",
    "創憶學坊(麗港城)",
    "創憶學坊(青衣)",
    "創憶學坊(港景峯)",
    "創憶學坊(鑽石山)"
])
start_date = st.date_input("Start Date", format="YYYY-MM-DD")
total_lessons = st.number_input("Total Number of Lessons", min_value=1, max_value=100, step=1)

frequency_options = list(weekday_map.keys())
selected_days = st.multiselect("Lesson Days of the Week", frequency_options)

if st.button("Generate Schedule"):
    if student_name and school_name and selected_days:
        schedule = generate_schedule(total_lessons, selected_days, start_date, student_name, school_name)
        st.success("Here is the schedule:")
        for line in schedule:
            st.write(line)
        
        # Calculate weeks range
        lessons_per_week = len(selected_days)
        total_weeks = calculate_weeks_range(total_lessons, lessons_per_week, [datetime.strptime(line.split('|')[1].strip().split(' ')[0] + ' ' + line.split('|')[1].strip().split(' ')[1] + ' ' + line.split('|')[1].strip().split(' ')[2], '%Y年 %m 月 %d 日').date() for line in schedule])
        
        end_date = start_date + timedelta(weeks=total_weeks) - timedelta(days=1)
        st.info(f"課程總期數範圍: {start_date} 至 {end_date} （共 {total_weeks} 週）")
    else:
        st.error("Please fill in all fields correctly.")

