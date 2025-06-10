import streamlit as st
from datetime import datetime, timedelta

# Weekday and Holiday Setup
weekday_map = {
    "Monday": 0,
    "Tuesday": 1,
    "Wednesday": 2,
    "Thursday": 3,
    "Friday": 4,
    "Saturday": 5,
    "Sunday": 6
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
st.title(":calendar: 課程日期安排 (淘大)")

student_name = st.text_input("Student Name")
school_name = "創憶學坊(淘大)"
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
    else:
        st.error("Please fill in all fields correctly.")

