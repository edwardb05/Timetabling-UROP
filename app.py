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
- Room assignments
- AEA requirements
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
    num_days = st.number_input("Number of Days for Exam Period", min_value=1, max_value=30, value=20)
    max_exams_2days = st.number_input("Maximum Exams in 2-Day Window", min_value=1, max_value=5, value=3)
    max_exams_5days = st.number_input("Maximum Exams in 5-Day Window", min_value=1, max_value=10, value=4)

with col2:
    week3_penalty = st.slider("Week 3 Penalty Weight", min_value=0, max_value=10, value=5)
    spread_penalty = st.slider("Exam Spacing Penalty Weight", min_value=0, max_value=10, value=5)
    extra_time_penalty = st.slider("Extra Time Student Penalty Weight", min_value=0, max_value=10, value=5)

# Fixed modules dictionary
Fixed_modules = {
    "BUSI60039 Business Strategy": [1,1],
    "BUSI60046 Project Management": [2,1],
    "ME-ELEC70098 Optimisation": [3,0],
    "MECH70001 Nuclear Thermal Hydraulics": [3,0],
    "BUSI60040/BUSI60043 Corporate Finance Online/Finance & Financial Management": [3,1],
    "MECH60004/MECH70042 Introduction to Nuclear Energy A/B": [4,0],
    "ME-ELEC70022 Modelling and Control of Multi-body Mechanical Systems": [4,0],
    "MATE97022 Nuclear Materials 1": [4,0],
    "ME-MATE70029 Nuclear Fusion": [9,0],
    "MECH70002 Nuclear Reactor Physics": [10,0],
    "ME-ELEC70076 Sustainable Electrical Systems": [10,0],
    "ME ELEC70066 Applied Advanced Optimisation": [10,0],
    "MECH70020 Combustion, Safety and Fire Dynamics": [11,0],
    "BIOE70016 Human Neuromechanical Control and Learning": [11,0],
    "CENG60013 Nuclear Chemical Engineering": [11,0],
    "MECH70008 Mechanical Transmissions Technology": [17,1],
    "MECH70006 Metal Processing Technology": [17,1],
    "MECH70021Aircraft Engine Technology": [17,1],
    "MECH70003 Future Clean Transport Technology": [17,1],
    "MECH60015/70030 PEN3/AME": [18,1]
}

# Core modules list
Core_modules = [
    "MECH70001 Nuclear Thermal Hydraulics",
    "MECH60004/MECH70042 Introduction to Nuclear Energy A/B",
    "MECH70002 Nuclear Reactor Physics",
    "MECH70008 Mechanical Transmissions Technology",
    "MECH70006 Metal Processing Technology",
    "MECH70021Aircraft Engine Technology",
    "MECH70003 Future Clean Transport Technology",
    "MECH60015/70030 PEN3/AME"
]

# Room dictionary with capacities and features
rooms = {
    'CAGB 203': [["Computer", "SEQ","AEA"], 65],
    'CAGB 309': [["SEQ"], 54],
    'CAGB 659-652': [["SEQ"], 75],
    'CAGB 747-748': [["SEQ"], 36],
    'CAGB 749-752': [["SEQ"], 75],
    'CAGB 761': [["Computer"], 25],
    'CAGB 762': [["Computer"], 25],
    'SKEM 208': [["Computer"], 35],
    'SKEM 317': [["Computer"], 20],
    'CAGB 320-321': [["AEA"], 10],
    'CAGB 305': [["AEA"], 4],
    'CAGB 349': [["AEA"], 2],
    'CAGB 311': [["AEA"], 1],
    'CAGB 765': [["AEA"], 10],
    'CAGB 527': [["AEA"], 2]
}

def ordinal(n):
    if 11 <= (n % 100) <= 13:
        return f"{n}th"
    else:
        return f"{n}{['th','st','nd','rd','th','th','th','th','th','th'][n % 10]}"

def process_files():
    if not all([student_file, module_file, dates_file]):
        st.error("Please upload all required files")
        return None, None, None, None, None, None, None, None, None

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

        # Create student-exam dictionary
        student_exams = {}
        student_rows = students_df.iloc[2:, :]
        
        # Get AEA students
        valid_aea_mask = (
            student_rows.iloc[:, 3].notna() &
            (student_rows.iloc[:, 3].astype(str).str.strip() != "#N/A")
        )
        AEA = student_rows.loc[valid_aea_mask, student_rows.columns[0]].tolist()

        for _, row in student_rows.iterrows():
            cid = row[0]
            exams_taken = []
            for col_idx, exam_name in enumerate(exams, start=9):
                if str(row[col_idx]).strip().lower() in ['x', 'a', 'b']:
                    exams_taken.append(exam_name)
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

        # Create days list with proper formatting
        days = []
        for i in range(21):
            date = first_monday + timedelta(days=i)
            day_str = date.strftime("%A ") + ordinal(date.day) + date.strftime(" %B")
            days.append(day_str)

        no_exam_dates = [[5,0],[5,1],[6,0],[6,1],[12,0],[12,1],[13,0],[13,1],[18,0],[19,0],[19,1],[20,0],[20,1]]
        for name, bh_date in bank_holidays:
            delta = (bh_date - first_monday).days
            if 0 <= delta <= 20:
                no_exam_dates.append([delta, 0])
                no_exam_dates.append([delta, 1])

        # Calculate exam counts for AEA and non-AEA students
        exam_counts = defaultdict(lambda: [0, 0])
        for cid, exams_taken in student_exams.items():
            if cid in AEA:
                for exam in exams_taken:
                    exam_counts[exam][0] += 1
            else:
                for exam in exams_taken:
                    exam_counts[exam][1] += 1

        return exams, student_exams, dict(leader_courses), no_exam_dates, extra_time_students_25, extra_time_students_50, AEA, exam_counts, days

    except Exception as e:
        st.error(f"Error processing files: {str(e)}")
        return None, None, None, None, None, None, None, None, None

def create_timetable(exams, student_exams, leader_courses, no_exam_dates, extra_time_students_25, extra_time_students_50, AEA, exam_counts, days):
    model = cp_model.CpModel()
    slots = [0, 1]  # 0 for morning, 1 for afternoon
    num_slots = len(slots)

    # Variables
    exam_day = {}
    exam_slot = {}
    for exam in exams:
        exam_day[exam] = model.NewIntVar(0, num_days - 1, f'{exam}_day')
        exam_slot[exam] = model.NewIntVar(0, num_slots - 1, f'{exam}_slot')

    # Room assignment variables
    exam_room = {}
    for exam in exams:
        for room in rooms:
            exam_room[(exam, room)] = model.NewBoolVar(f'{exam}_in_{room.replace(" ", "_")}')

    # Constraint 1: Core modules can't be on the same day as other modules
    for student, exs in student_exams.items():
        core_mods = [exam for exam in exs if exam in Core_modules]
        other_mods = [exam for exam in exs if exam not in Core_modules]
        for exam in core_mods:
            for other in other_mods:
                model.Add(exam_day[exam] != exam_day[other])

    # Constraint 2: Fixed module dates
    for exam, (day_fixed, slot_fixed) in Fixed_modules.items():
        if exam in exams:
            model.Add(exam_day[exam] == day_fixed)
            model.Add(exam_slot[exam] == slot_fixed)

    # Constraint 3: Forbidden exam day-slot assignments
    for exam in exams:
        for day, slot in no_exam_dates:
            model.AddForbiddenAssignments([exam_day[exam], exam_slot[exam]], [(day, slot)])

    # Constraint 4: Max 3 exams in any 2-day window per student
    for student, exs in student_exams.items():
        for d in range(num_days - 1):
            exams_in_2_days = []
            for exam in exs:
                is_on_d = model.NewBoolVar(f'{student}_{exam}_on_day_{d}')
                is_on_d1 = model.NewBoolVar(f'{student}_{exam}_on_day_{d+1}')
                is_on_either = model.NewBoolVar(f'{student}_{exam}_on_day_{d}_or_{d+1}')

                model.Add(exam_day[exam] == d).OnlyEnforceIf(is_on_d)
                model.Add(exam_day[exam] != d).OnlyEnforceIf(is_on_d.Not())
                model.Add(exam_day[exam] == d + 1).OnlyEnforceIf(is_on_d1)
                model.Add(exam_day[exam] != d + 1).OnlyEnforceIf(is_on_d1.Not())
                model.AddBoolOr([is_on_d, is_on_d1]).OnlyEnforceIf(is_on_either)
                model.AddBoolAnd([is_on_d.Not(), is_on_d1.Not()]).OnlyEnforceIf(is_on_either.Not())

                exams_in_2_days.append(is_on_either)

            model.Add(sum(exams_in_2_days) <= max_exams_2days)

    # Constraint 5: Max 4 exams in any 5-day window per student
    for student, exs in student_exams.items():
        for start_day in range(num_days - 4):
            exams_in_window = []
            for exam in exs:
                in_window = model.NewBoolVar(f'{student}_{exam}_in_day_{start_day}_to_{start_day + 4}')
                model.AddLinearConstraint(exam_day[exam], start_day, start_day + 4).OnlyEnforceIf(in_window)
                before_window = model.NewBoolVar(f'{student}_{exam}_before_{start_day}')
                after_window = model.NewBoolVar(f'{student}_{exam}_after_{start_day + 4}')
                model.Add(exam_day[exam] < start_day).OnlyEnforceIf(before_window)
                model.Add(exam_day[exam] >= start_day).OnlyEnforceIf(before_window.Not())
                model.Add(exam_day[exam] > start_day + 4).OnlyEnforceIf(after_window)
                model.Add(exam_day[exam] <= start_day + 4).OnlyEnforceIf(after_window.Not())
                model.AddBoolOr([before_window, after_window]).OnlyEnforceIf(in_window.Not())
                exams_in_window.append(in_window)
            model.Add(sum(exams_in_window) <= max_exams_5days)

    # Constraint 6: At most 1 exam in week 3 per module leader
    for leader, leader_exams in leader_courses.items():
        exams_in_week3 = []
        for exam in leader_exams:
            in_week3 = model.NewBoolVar(f'{exam}_in_week3')
            model.AddLinearConstraint(exam_day[exam], 13, 20).OnlyEnforceIf(in_week3)
            before_week3 = model.NewBoolVar(f'{exam}_before_week3')
            after_week3 = model.NewBoolVar(f'{exam}_after_week3')
            model.Add(exam_day[exam] < 13).OnlyEnforceIf(before_week3)
            model.Add(exam_day[exam] >= 13).OnlyEnforceIf(before_week3.Not())
            model.Add(exam_day[exam] > 20).OnlyEnforceIf(after_week3)
            model.Add(exam_day[exam] <= 20).OnlyEnforceIf(after_week3.Not())
            model.AddBoolOr([before_week3, after_week3]).OnlyEnforceIf(in_week3.Not())
            exams_in_week3.append(in_week3)
        model.Add(sum(exams_in_week3) <= 1)

    # Constraint 7: Extra time 50% students: max 1 exam per day
    for student in extra_time_students_50:
        for day in range(num_days):
            exams_on_day = []
            for exam in student_exams[student]:
                is_on_day = model.NewBoolVar(f'{student}_{exam}_on_day_{day}')
                model.Add(exam_day[exam] == day).OnlyEnforceIf(is_on_day)
                model.Add(exam_day[exam] != day).OnlyEnforceIf(is_on_day.Not())
                exams_on_day.append(is_on_day)
            model.Add(sum(exams_on_day) <= 1)

    # Constraint 8: Room assignments
    for exam in exams:
        # Each exam must be assigned to at least one room
        model.Add(sum(exam_room[(exam, room)] for room in rooms) >= 1)
        
        # Room capacity constraints
        for room, (features, capacity) in rooms.items():
            # Get number of AEA and non-AEA students for this exam
            aea_count, non_aea_count = exam_counts[exam]
            
            # If room has AEA feature, it can take AEA students
            if "AEA" in features:
                model.Add(aea_count * exam_room[(exam, room)] <= capacity)
            else:
                # If room doesn't have AEA feature, it can't be used for AEA students
                model.Add(aea_count * exam_room[(exam, room)] == 0)
            
            # Total capacity constraint
            model.Add((aea_count + non_aea_count) * exam_room[(exam, room)] <= capacity)

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

            num_exams = model.NewIntVar(0, len(exams_on_day), f'{student}_num_exams_day_{day}')
            model.Add(num_exams == sum(exams_on_day))

            has_multiple_exams = model.NewBoolVar(f'{student}_more_than_one_exam_day_{day}')
            model.Add(num_exams >= 2).OnlyEnforceIf(has_multiple_exams)
            model.Add(num_exams < 2).OnlyEnforceIf(has_multiple_exams.Not())

            penalty = model.NewIntVar(0, 1, f'{student}_penalty_day_{day}')
            model.Add(penalty == 1).OnlyEnforceIf(has_multiple_exams)
            model.Add(penalty == 0).OnlyEnforceIf(has_multiple_exams.Not())

            soft_penalties.append(extra_time_penalty * penalty)

    # Spread penalties for module leaders
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
                spread_penalties.append(spread_penalty * close_penalty)

    # Minimize total penalties
    model.Minimize(sum(spread_penalties) + sum(soft_penalties))

    # Solve
    solver = cp_model.CpSolver()
    status = solver.Solve(model)

    if status == cp_model.FEASIBLE or status == cp_model.OPTIMAL:
        timetable = {}
        room_assignments = {}
        for exam in exams:
            d = solver.Value(exam_day[exam])
            s = solver.Value(exam_slot[exam])
            timetable[exam] = (d, s)
            
            # Get room assignments
            assigned_rooms = [room for room in rooms if solver.Value(exam_room[(exam, room)])]
            room_assignments[exam] = assigned_rooms
            
        return timetable, room_assignments
    else:
        return None, None

def visualize_timetable(timetable, room_assignments, days):
    if not timetable:
        return None

    # Create a DataFrame for visualization
    data = []
    for exam, (day, slot) in timetable.items():
        rooms_str = ", ".join(room_assignments[exam])
        data.append({
            'Exam': exam,
            'Day': days[day],
            'Slot': 'Morning' if slot == 0 else 'Afternoon',
            'Rooms': rooms_str
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
    
    return fig, df

if st.button("Generate Timetable"):
    with st.spinner("Processing files and generating timetable..."):
        result = process_files()
        if all(result):
            exams, student_exams, leader_courses, no_exam_dates, extra_time_students_25, extra_time_students_50, AEA, exam_counts, days = result
            timetable, room_assignments = create_timetable(exams, student_exams, leader_courses, no_exam_dates, extra_time_students_25, extra_time_students_50, AEA, exam_counts, days)
            
            if timetable:
                st.success("Timetable generated successfully!")
                
                # Display timetable
                st.header("Generated Timetable")
                fig, df = visualize_timetable(timetable, room_assignments, days)
                if fig:
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # Display detailed timetable
                    st.subheader("Detailed Timetable")
                    st.dataframe(df)
                
                # Export option
                if st.button("Export Timetable"):
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