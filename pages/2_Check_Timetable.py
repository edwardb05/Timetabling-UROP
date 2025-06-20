# Timetable Checking Page


import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from collections import defaultdict
import pickle


try:
    with open("timetable_data.pkl", "rb") as f:
        data = pickle.load(f)

    # Unpack all variables
    days = data["days"]
    slots = data["slots"]
    exams = data["exams"]
    AEA = data["AEA"]
    leader_courses = data["leader_courses"]
    extra_time_students_25 = data["extra_time_students_25"]
    extra_time_students_50 = data["extra_time_students_50"]
    student_exams = data["student_exams"]
    exam_counts = data["exam_counts"]
    Fixed_modules = data["Fixed_modules"]
    Core_modules = data["Core_modules"]
    rooms = data["rooms"]
    exam_types = data["exam_types"]

    st.success("Timetable data loaded successfully from file!")

except FileNotFoundError:
    st.warning("Timetable file not found. Please generate a timetable first.")
except KeyError as e:
    st.error(f"Missing key in saved data: {e}")
except Exception as e:
    st.error(f"Failed to load timetable data: {e}")
    def file_reading(filepath, days, slots):
        df = pd.read_excel(filepath)
        exams_timetabled = {}

        for _, row in df.iterrows():
            exam_name = row['Exam']


            day_name = day_name if pd.isna(row['Date']) else row['Date']
            slot_name = slot_name if pd.isna(row['Time']) else (0 if row['Time'] == "Morning" else 1)
            if pd.isna(exam_name) or exam_name == '':
                continue  # Skip empty rows
            room = row['Room'].split(', ') if pd.notna(row['Room']) and row['Room'] else []

            try:
                d = days.index(day_name)
                s = slots.index(slot_name)
            except ValueError:
                raise ValueError(f"Unrecognized day or slot in file: {day_name} / {slot_name}")

            exams_timetabled[exam_name] = (d, s, room)

        return exams_timetabled

    def file_checking(exams_timetabled, Fixed_modules, Core_modules, student_exams, leader_courses, extra_time_students_50, exams, AEA,exam_counts):
        def get_full_schedule(exams_timetabled, Fixed_modules):
            full_schedule = Fixed_modules.copy()
            full_schedule.update(exams_timetabled)
            return full_schedule
        def check_exam_constraints(student_exams, exams_timetabled, Fixed_modules, Core_modules, module_leaders, extra_time_students_50, exams,AEA,):
            violations = []
            schedule = get_full_schedule(exams_timetabled, Fixed_modules)
            for exam in exams:
                if exam not in schedule:
                    violations.append(f"❌ Exam '{exam}' is not scheduled in the timetable.")


            for student, exams in student_exams.items():
                day_count = defaultdict(int)

                for exam in exams:
                    if exam in schedule:
                        day = schedule[exam][0]
                        day_count[day] += 1
                days = sorted(day_count.keys())
                for day in days:
                    next_day = day + 1
                    if next_day in day_count:
                        total = day_count[day] + day_count[next_day]
                        if total > 3:
                            violations.append(
                                f"❌ Student {student} has more than 3 exams across days {day} and {next_day}"
                            )

            for student, exams in student_exams.items():
                day_count = defaultdict(int)
                for exam in exams:
                    if exam in schedule:
                        day = schedule[exam][0]
                        day_count[day] += 1
            all_days = sorted(day_count.keys())
            if all_days:
                min_day, max_day = all_days[0], all_days[-1]
                for start_day in range(min_day, max_day - 4 + 1):
                    total = sum(day_count.get(day, 0) for day in range(start_day, start_day + 5))
                    if total > 4:
                        violations.append(
                            f"❌ Student {student} has more than 4 exams from day {start_day} to {start_day + 4}"
                        )
            for student, exams in student_exams.items():
                core_mods = [exam for exam in exams if exam in Core_modules]
                other_mods = [exam for exam in exams if exam not in Core_modules]
                for core_exam in core_mods:
                    core_day = exams_timetabled[core_exam][0]
                    for other_exam in other_mods:
                        other_day = exams_timetabled[other_exam][0]
                        if core_day == other_day:
                            violations.append(
                                f"❌ Student {student} has core exam '{core_exam}' and non-core exam '{other_exam}' on the same day ({core_day})"
                            )
            week3_days = set(range(15, 21))
            for leader, mods in module_leaders.items():
                exams_in_week3 = [exam for exam in mods if exam in schedule and schedule[exam][0] in week3_days]
                if len(exams_in_week3) > 1:
                    violations.append(f"❌ Module leader {leader} has more than one exam in week 3: {exams_in_week3}")
            for student in extra_time_students_50:
                if student not in student_exams:
                    continue
                day_count = defaultdict(int)
                for exam in student_exams[student]:
                    if exam in schedule:
                        day = schedule[exam][0]
                        day_count[day] += 1
                for day, count in day_count.items():
                    if count > 1:
                        violations.append(f"❌ Student {student} with >50% extra time has {count} exams on day {day}")
            for student in AEA:
                if student not in extra_time_students_50:
                    day_count = defaultdict(int)
                    for exam in student_exams[student]:
                        if exam in schedule:
                            day = schedule[exam][0]
                            day_count[day] += 1
                    for day, count in day_count.items():
                        if count > 1:
                            violations.append(f"⚠️soft warning Student {student} with <=25% extra time has {count} exams on day {day}")
            return violations
        def check_room_constraints(
            exams_timetabled,      # dict: exam -> (day, slot, [assigned_rooms])
            exam_counts,           # dict: exam -> (AEA_students, SEQ_students)
            room_dict,
            exam_types,              # dict: room_name -> [list of types, capacity]
        ):
            violations = []
            room_schedule = defaultdict(list)
            for exam, (day, slot, rooms_) in exams_timetabled.items():
                for room in rooms_:
                    room_schedule[(day, slot, room)].append(exam)
            for (day, slot, room), exams_in_room in room_schedule.items():
                if len(exams_in_room) > 1:
                    violations.append(
                        f"❌ Room '{room}' double-booked on day {day}, slot {slot} for exams: {exams_in_room}"
                    )

            for exam, (day, slot, rooms) in exams_timetabled.items():
                if not rooms:
                    violations.append(f"❌ Exam '{exam}' has no assigned room!")
            for exam, (day, slot, rooms) in exams_timetabled.items():
                if exam not in exam_counts:
                    violations.append(f"⚠️ No student count for exam '{exam}', skipping capacity check")
                    continue
                AEA_students, SEQ_students = exam_counts[exam]
                AEA_capacity = sum(room_dict[r][1] for r in rooms if "AEA" in room_dict[r][0])
                SEQ_capacity = sum(room_dict[r][1] for r in rooms if "SEQ" in room_dict[r][0])
                if AEA_capacity < AEA_students:
                    violations.append(
                        f"❌ Exam '{exam}' has insufficient AEA capacity: needed {AEA_students}, assigned {AEA_capacity}"
                    )
                if SEQ_capacity < SEQ_students:
                    violations.append(
                        f"❌ Exam '{exam}' has insufficient SEQ capacity: needed {SEQ_students}, assigned {SEQ_capacity}"
                    )
            for exam, (day, slot, rooms) in exams_timetabled.items():
                if exam_types[exam] == "PC":
                    for room in rooms:
                        if "Computer" not in room_dict[room][0]:
                            violations.append(
                                f"❌ Computer-based exam '{exam}' assigned to non-computer room '{room}'"
                            )
            return violations
        violations = check_exam_constraints(
            student_exams=student_exams,
            exams_timetabled=exams_timetabled,
            Fixed_modules=Fixed_modules,
            Core_modules=Core_modules,
            module_leaders=leader_courses,
            extra_time_students_50=extra_time_students_50,
            exams = exams,
            AEA = AEA
        )

        violations.extend(check_room_constraints(
            exams_timetabled=exams_timetabled,
            exam_counts=exam_counts,
            room_dict=rooms,
            exam_types=exam_types
        ))
        if violations:
            for v in violations:
                st.write(v)
        else:
            st.write("✅ All constraints satisfied! No violations found.")

    uploaded_file = st.file_uploader("Upload a file to check", type=["xlsx", "csv"])

    if st.button("🔍 Check Files"):
        st.header("🔍 Check Your Files")
        if uploaded_file is not None:
            try:
                st.write("✅ File uploaded successfully!")
                exams_timetabled = file_reading(uploaded_file, days, slots)
                file_checking(exams_timetabled, Fixed_modules, Core_modules, student_exams, leader_courses, extra_time_students_50, exams, AEA,exam_counts)
            except Exception as e:
                st.error(f"Error reading file: {e}") 