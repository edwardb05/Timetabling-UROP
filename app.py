import streamlit as st
import pandas as pd
import numpy as np
from ortools.sat.python import cp_model
from rapidfuzz import process, fuzz
from collections import defaultdict
from openpyxl import load_workbook
from datetime import datetime, timedelta
import re
from dateutil.parser import parse
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Exam Timetabling System", layout="wide")

st.title("Exam Timetabling System")
st.markdown("""
This app helps create an optimal exam timetable while considering various constraints like:
- Student exam conflicts
- Module leader preferences
- Special accommodations for students with extra time
- Fixed exam dates
- Bank holidays
""")

# File upload section
st.header("Upload Required Files")
col1, col2, col3 = st.columns(3)

with col1:
    student_file = st.file_uploader("Upload Student List", type=['xlsx'])
with col2:
    module_file = st.file_uploader("Upload Module List", type=['xlsx'])
with col3:
    dates_file = st.file_uploader("Upload Useful Dates", type=['xlsx'])

# Parameters section
st.header("Timetabling Parameters")
col1, col2 = st.columns(2)

with col1:
    num_days = st.number_input("Number of Days for Exam Period", min_value=1, max_value=30, value=21)
    max_exams_2days = st.number_input("Maximum Exams in 2-Day Window", min_value=1, max_value=5, value=3)
    max_exams_5days = st.number_input("Maximum Exams in 5-Day Window", min_value=1, max_value=10, value=4)

with col2:
    week3_penalty = st.slider("Week 3 Penalty Weight", min_value=0, max_value=10, value=5)
    spread_penalty = st.slider("Exam Spacing Penalty Weight", min_value=0, max_value=10, value=5)

def process_files():
    if not all([student_file, module_file, dates_file]):
        st.error("Please upload all required files")
        return None, None, None, None, None, None

    try:
        # Read student data
        students_df = pd.read_excel(student_file, header=None)
        
        # Read module data
        leaders_df = pd.read_excel(module_file, sheet_name=1, header=1)
        
        # Read dates
        wb = load_workbook(dates_file)
        ws = wb.active

        # Extract exams from student data
        exams = students_df.iloc[0, 9:].dropna().tolist()
        exams.append('MECH60015 / MECH70030 Professional Engineering Skills (ME3/AME) (ECE exam)')

        # Create student-exam dictionary
        student_exams = {}
        student_rows = students_df.iloc[2:, :]
        for _, row in student_rows.iterrows():
            cid = row[0]
            exams_taken = []
            for col_idx, exam_name in enumerate(exams[:-1], start=9):
                if str(row[col_idx]).strip().lower() == 'x':
                    exams_taken.append(exam_name)
            exams_taken.append('MECH60015 / MECH70030 Professional Engineering Skills (ME3/AME) (ECE exam)')
            student_exams[cid] = exams_taken

        # Process module leaders
        standardized_names = exams
        leader_courses = defaultdict(list)
        for _, row in leaders_df.iterrows():
            leader = row['Module Leader (lecturer 1)']
            name = row['Module Name']
            code = row['Banner Code (New CR)']

            if pd.isna(code) or pd.isna(name) or pd.isna(leader) or leader == "n/a":
                continue

            combined_name = f"{code} {name}"
            best_match, score, _ = process.extractOne(
                combined_name, standardized_names, scorer=fuzz.token_sort_ratio
            )

            if score >= 70 and best_match not in leader_courses[leader]:
                leader_courses[leader].append(best_match)

        # Process extra time students
        extra_time_students_25 = students_df[students_df.iloc[:, 3].astype(str).str.startswith(("15min/hour", "25% extra time"))].iloc[:, 0].tolist()
        extra_time_students_50 = students_df[students_df.iloc[:, 3].astype(str).str.startswith(("30min/hour", "50% extra time"))].iloc[:, 0].tolist()

        # Process bank holidays
        bank_holidays = []
        row = 5
        while True:
            name = ws[f"F{row}"].value
            date_cell = ws[f"G{row}"].value
            if name is None or "Term Dates" in str(name):
                break
            if isinstance(date_cell, datetime):
                bank_holidays.append((str(name).strip(), date_cell.date()))
            row += 1

        # Find Summer Term start date
        summer_start = None
        while row < ws.max_row:
            cell_value = ws[f"F{row}"].value
            if cell_value and "Summer Term" in str(cell_value):
                term_range = ws[f"F{row + 1}"].value
                if term_range:
                    start_part = term_range.split("to")[0].strip()
                    start_str = re.sub(r"^\w+\s+", "", start_part)
                    year_match = re.search(r"\b\d{4}\b", term_range)
                    if year_match:
                        start_str += f" {year_match.group(0)}"
                    summer_start = parse(start_str, dayfirst=True).date()
                break
            row += 1

        if not summer_start:
            raise ValueError("Summer Term start date not found")

        first_monday = summer_start
        while first_monday.weekday() != 0:
            first_monday += timedelta(days=1)

        no_exam_dates = [[5,0],[5,1],[6,0],[6,1],[12,0],[12,1],[13,0],[13,1],[20,0]]
        for name, bh_date in bank_holidays:
            delta = (bh_date - first_monday).days
            if 0 <= delta <= 20:
                no_exam_dates.append([delta, 0])
                no_exam_dates.append([delta, 1])

        return exams, student_exams, dict(leader_courses), no_exam_dates, extra_time_students_25, extra_time_students_50

    except Exception as e:
        st.error(f"Error processing files: {str(e)}")
        return None, None, None, None, None, None

def create_timetable(exams, student_exams, leader_courses, no_exam_dates, extra_time_students_25, extra_time_students_50):
    model = cp_model.CpModel()
    slots = [0, 1]  # 0 for morning, 1 for afternoon
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"] * 3
    num_slots = len(slots)

    # Variables
    exam_day = {}
    exam_slot = {}
    for exam in exams:
        exam_day[exam] = model.NewIntVar(0, num_days - 1, f'{exam}_day')
        exam_slot[exam] = model.NewIntVar(0, num_slots - 1, f'{exam}_slot')

    # Constraint 1: No more than 3 exams in a 2-day window
    for student in student_exams:
        for d in range(num_days - 1):
            exams_in_2_days = []
            for exam in student_exams[student]:
                is_on_d = model.NewBoolVar(f'{student}_{exam}_on_day_{d}')
                is_on_d1 = model.NewBoolVar(f'{student}_{exam}_on_day_{d+1}')
                is_in_either = model.NewBoolVar(f'{student}_{exam}_on_day_{d}_or_{d+1}')

                model.Add(exam_day[exam] == d).OnlyEnforceIf(is_on_d)
                model.Add(exam_day[exam] != d).OnlyEnforceIf(is_on_d.Not())
                model.Add(exam_day[exam] == d+1).OnlyEnforceIf(is_on_d1)
                model.Add(exam_day[exam] != d+1).OnlyEnforceIf(is_on_d1.Not())

                model.AddBoolOr([is_on_d, is_on_d1]).OnlyEnforceIf(is_in_either)
                model.AddBoolAnd([is_on_d.Not(), is_on_d1.Not()]).OnlyEnforceIf(is_in_either.Not())

                exams_in_2_days.append(is_in_either)

            model.Add(sum(exams_in_2_days) <= max_exams_2days)

    # Constraint 2: No more than 4 exams in a 5-day window
    for student in student_exams:
        for start_day in range(num_days - 4):
            exams_in_window = []
            for exam in student_exams[student]:
                in_window = model.NewBoolVar(f'{student}_{exam}_in_day_{start_day}_to_{start_day+4}')
                model.AddLinearConstraint(exam_day[exam], start_day, start_day + 4).OnlyEnforceIf(in_window)
                exams_in_window.append(in_window)
            model.Add(sum(exams_in_window) <= max_exams_5days)

    # Constraint 3: Core modules can't be on the same day as other modules
    Core_modules = ["MECH70001 Nuclear Thermal Hydraulics", "MECH60004/MECH70042 Introduction to Nuclear Energy A/B",
                   "MECH70002 Nuclear Reactor Physics", "MECH70008 Mechanical Transmissions Technology",
                   "MECH70006 Metal Processing Technology", "MECH70021Aircraft Engine Technology",
                   "MECH70003 Future Clean Transport Technology",
                   "MECH60015 / MECH70030 Professional Engineering Skills (ME3/AME) (ECE exam)"]

    for student, exams_list in student_exams.items():
        core_mods = [exam for exam in exams_list if exam in Core_modules]
        other_mods = [exam for exam in exams_list if exam not in Core_modules]
        for exam in core_mods:
            for other in other_mods:
                model.Add(exam_day[exam] != exam_day[other])

    # Constraint 4: Module leaders can have at most one exam in week 3
    for leader in leader_courses:
        week_3_exams = []
        for exam in leader_courses[leader]:
            is_in_week3 = model.NewBoolVar(f'{exam}_in_week3')
            model.AddLinearConstraint(exam_day[exam], 13, 20).OnlyEnforceIf(is_in_week3)
            week_3_exams.append(is_in_week3)
        model.Add(sum(week_3_exams) <= 1)

    # Constraint 5: Students with 50% extra time can't have more than one exam per day
    for student in extra_time_students_50:
        for day in range(num_days):
            exams_on_day = []
            for exam in student_exams[student]:
                is_on_day = model.NewBoolVar(f'{student}_{exam}_on_day_{day}')
                model.Add(exam_day[exam] == day).OnlyEnforceIf(is_on_day)
                model.Add(exam_day[exam] != day).OnlyEnforceIf(is_on_day.Not())
                exams_on_day.append(is_on_day)
            model.Add(sum(exams_on_day) <= 1)

    # Constraint 6: No exams on forbidden dates
    for exam in exams:
        for date, slot in no_exam_dates:
            is_forbidden = model.NewBoolVar(f'{exam}_forbidden_{date}_{slot}')
            model.Add(exam_day[exam] == date).OnlyEnforceIf(is_forbidden)
            model.Add(exam_day[exam] != date).OnlyEnforceIf(is_forbidden.Not())
            model.Add(exam_slot[exam] == slot).OnlyEnforceIf(is_forbidden)
            model.Add(exam_slot[exam] != slot).OnlyEnforceIf(is_forbidden.Not())
            model.Add(is_forbidden == 0)

    # Soft constraints and penalties
    soft_penalties = []
    for student in extra_time_students_25:
        for day in range(num_days):
            exams_on_day = []
            for exam in student_exams[student]:
                is_on_day = model.NewBoolVar(f'{student}_{exam}_on_day_{day}')
                model.Add(exam_day[exam] == day).OnlyEnforceIf(is_on_day)
                model.Add(exam_day[exam] != day).OnlyEnforceIf(is_on_day.Not())
                exams_on_day.append(is_on_day)
            penalty = model.NewIntVar(0, 1, f'{student}_penalty_day_{day}')
            model.Add(penalty == 1).OnlyEnforceIf(len(exams_on_day) > 1)
            model.Add(penalty == 0).OnlyEnforceIf(len(exams_on_day) <= 1)
            soft_penalties.append(penalty)

    # Penalties for exams being too close together
    spread_penalties = []
    for leader in leader_courses:
        mods = leader_courses[leader]
        for i in range(len(mods)):
            for j in range(i+1, len(mods)):
                m1 = mods[i]
                m2 = mods[j]
                diff = model.NewIntVar(-21, 21, f'{m1}_{m2}_diff')
                abs_diff = model.NewIntVar(0, 21, f'{m1}_{m2}_abs_diff')
                model.Add(diff == exam_day[m1] - exam_day[m2])
                model.AddAbsEquality(abs_diff, diff)

                close_penalty = model.NewIntVar(0, 5, f'{m1}_{m2}_penalty')
                is_gap_3 = model.NewBoolVar(f'{m1}_{m2}_gap3')
                is_gap_2 = model.NewBoolVar(f'{m1}_{m2}_gap2')
                is_gap_1 = model.NewBoolVar(f'{m1}_{m2}_gap1')
                is_gap_0 = model.NewBoolVar(f'{m1}_{m2}_gap0')

                model.Add(abs_diff == 3).OnlyEnforceIf(is_gap_3)
                model.Add(abs_diff != 3).OnlyEnforceIf(is_gap_3.Not())
                model.Add(abs_diff == 2).OnlyEnforceIf(is_gap_2)
                model.Add(abs_diff != 2).OnlyEnforceIf(is_gap_2.Not())
                model.Add(abs_diff == 1).OnlyEnforceIf(is_gap_1)
                model.Add(abs_diff != 1).OnlyEnforceIf(is_gap_1.Not())
                model.Add(abs_diff == 0).OnlyEnforceIf(is_gap_0)
                model.Add(abs_diff != 0).OnlyEnforceIf(is_gap_0.Not())

                model.Add(close_penalty == 1).OnlyEnforceIf(is_gap_3)
                model.Add(close_penalty == 3).OnlyEnforceIf(is_gap_2)
                model.Add(close_penalty == 4).OnlyEnforceIf(is_gap_1)
                model.Add(close_penalty == 5).OnlyEnforceIf(is_gap_0)
                model.Add(close_penalty == 0).OnlyEnforceIf(
                    is_gap_3.Not(), is_gap_2.Not(), is_gap_1.Not(), is_gap_0.Not()
                )
                spread_penalties.append(close_penalty)

    # Minimize total penalties
    model.Minimize(sum(spread_penalties) * spread_penalty + sum(soft_penalties) * week3_penalty)

    # Solve
    solver = cp_model.CpSolver()
    status = solver.Solve(model)

    if status == cp_model.FEASIBLE or status == cp_model.OPTIMAL:
        timetable = {}
        for exam in exams:
            d = solver.Value(exam_day[exam])
            s = solver.Value(exam_slot[exam])
            timetable[exam] = (d, s)
        return timetable
    else:
        return None

def visualize_timetable(timetable, days):
    if not timetable:
        return None

    # Create a DataFrame for visualization
    data = []
    for exam, (day, slot) in timetable.items():
        data.append({
            'Exam': exam,
            'Day': days[day],
            'Slot': 'Morning' if slot == 0 else 'Afternoon'
        })
    
    df = pd.DataFrame(data)
    
    # Create a heatmap
    pivot = pd.crosstab(df['Day'], df['Slot'])
    
    fig = go.Figure(data=go.Heatmap(
        z=pivot.values,
        x=pivot.columns,
        y=pivot.index,
        colorscale='Viridis'
    ))
    
    fig.update_layout(
        title='Exam Timetable Heatmap',
        xaxis_title='Time Slot',
        yaxis_title='Day'
    )
    
    return fig

if st.button("Generate Timetable"):
    with st.spinner("Processing files and generating timetable..."):
        exams, student_exams, leader_courses, no_exam_dates, extra_time_students_25, extra_time_students_50 = process_files()
        
        if all([exams, student_exams, leader_courses, no_exam_dates]):
            timetable = create_timetable(exams, student_exams, leader_courses, no_exam_dates, extra_time_students_25, extra_time_students_50)
            
            if timetable:
                st.success("Timetable generated successfully!")
                
                # Display timetable
                st.header("Generated Timetable")
                fig = visualize_timetable(timetable, days)
                if fig:
                    st.plotly_chart(fig, use_container_width=True)
                
                # Export option
                if st.button("Export Timetable"):
                    df = pd.DataFrame([
                        {'Exam': exam, 'Day': days[day], 'Slot': 'Morning' if slot == 0 else 'Afternoon'}
                        for exam, (day, slot) in timetable.items()
                    ])
                    csv = df.to_csv(index=False)
                    st.download_button(
                        "Download CSV",
                        csv,
                        "exam_timetable.csv",
                        "text/csv",
                        key='download-csv'
                    )
            else:
                st.error("Could not generate a valid timetable. Please check the constraints.") 