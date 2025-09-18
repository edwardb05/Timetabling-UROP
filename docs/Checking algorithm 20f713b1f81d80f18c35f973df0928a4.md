# Checking algorithm

Owner: Edward Brady

Once a time table is created it can be checked to ensure that it reaches all conditions, this can also be used to feed an excel spreadsheet in to check that the exams in it fit the constraints. These should be similar to setting the constraints but a little less confusing as it can use simple python logic rather than using the OR-tools model logic.

- Starts by extracting useful data from an excel spreadsheet, this is to allow an updated version to be uploaded.
    
    ```python
    def load_exam_schedule(filepath, days, slots):
        """
        Reconstructs the exams_timetabled dictionary from a saved exam schedule Excel file.
    
        Parameters:
            filepath (str): Path to the Excel file.
            days (list): Ordered list of day names used in the original schedule.
            slots (list): ['Morning', 'Afternoon']
    
        Returns:
            dict: exams_timetabled with structure {exam_name: (day_index, slot_index, [rooms])}
        """
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
    ```
    
    Days and slots must be initialised in the generation of a timetable thus this must happen after.
    
    As the days and times are merged this uses the same day from before if the row is N/A same for time . This produces a dictionary same as the one produced by the solver.
    
- Runs two modules and collects all the violations from student exams and room constraints
    
    ```python
     violations = check_exam_constraints(
          student_exams=student_exams,
          exams_timetabled=exams_timetabled,
          Fixed_modules=Fixed_modules,
          Core_modules=Core_modules,
          module_leaders=leader_courses,
          extra_time_students_50=extra_time_students_50,
          exams = exams,
      )
    
      violations.extend(check_room_constraints(
          exams_timetabled=exams_timetabled,
          exam_counts=exam_counts,
          room_dict=rooms
      ))
    ```
    

## Student Exam Constraints

- Checks all exams have been scheduled correctly
    
    ```python
    for exam in exams:
    	if exam not in schedule:
    		violations.append(f"❌ Exam '{exam}' is not scheduled in the timetable.")
    ```
    
- Check no students have two exams in the same slut
    
    ```python
        # 0. Students can't have two exams at the same time
        for student, exams in student_exams.items():
            for i in range(len(exams)):
                for j in range(i + 1, len(exams)):
                    exam1 = exams[i]
                    exam2 = exams[j]
                    if exams_timetabled[exam1][0] == exams_timetabled[exam2][0] and exams_timetabled[exam1][1] == exams_timetabled[exam2][1]:
                        violations.append(
                            f"❌ Student {student} has two exams '{exam1}' and '{exam2}' at the same time"
                        )
        
    ```
    
- Core modules are fixed at students cannot have more than one exam on this day (In red on timetable)
    
    ```python
        for student, exams in student_exams.items():
            core_mods = [exam for exam in exams if exam in Core_modules]
            other_mods = [exam for exam in exams if exam not in Core_modules]
    
            for core_exam in core_mods:
                core_day = exams_timetabled[core_exam][0]  # Assume (day, slot, rooms)
    
                for other_exam in other_mods:
                    other_day = exams_timetabled[other_exam][0]
    
                    if core_day == other_day:
                        violations.append(
                            f"❌ Student {student} has core exam '{core_exam}' and non-core exam '{other_exam}' on the same day ({core_day})"
                            )
    ```
    
    This loops through each student and gets a list of all their core exams and non core exams it then checks if the day of any of their core exams is the same as any other exams
    
- Other modules are fixed in date and time (In yellow on timetable )
    
    
    ```python
        for exam, fixed_slot in Fixed_modules.items():
            scheduled_slot = exams_timetabled.get(exam)
            if scheduled_slot != fixed_slot:
                violations.append(f"❌ Fixed module '{exam}' is not at the correct time (expected {fixed_slot}, got {scheduled_slot}).")
    
    ```
    
    Filters through each exam in the list of fixed ones and ensures they are on the correct days.
    
- No more than three exams in two days
    
    ```python
    # 3. No more than 3 exams in any 2 consecutive days (per student)
        for student, exams in student_exams.items():
            day_count = defaultdict(int)
            for exam in exams:
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
    
    ```
    
    This makes a dictionary of all the days and amount of exams on that day and then loops through to see if two consecutive days have more than 3 exams
    
- No more than four exams in five consecutive days (Monday to Friday)
    
    ```python
    # 4. No more than 4 exams in any 5 consecutive weekdays (Monday to Friday)
        for student, exams in student_exams.items():
            day_count = defaultdict(int)
            for exam in exams:
                if exam in schedule:
                    day = schedule[exam][0]
                    day_count[day] += 1
    
        all_days = sorted(day_count.keys())
        if all_days:
            min_day, max_day = all_days[0], all_days[-1]
            # Slide over every possible consecutive 5-day window in the exam period
            for start_day in range(min_day, max_day - 4 + 1):
                total = sum(day_count.get(day, 0) for day in range(start_day, start_day + 5))
                if total > 4:
                    violations.append(
                        f"❌ Student {student} has more than 4 exams from day {start_day} to {start_day + 4}"
                    )
    
    ```
    
    Same as the constraint above this initaites a day count and then sums it up for each 5 day window and check it always falls below 4
    
- Module leaders can’t have more than one exam in the third week
    
    ```python
      # 5. Module leaders cannot have more than one exam in the third week (days 15 to 20 inclusive)
        week3_days = set(range(15, 21))
        for leader, mods in module_leaders.items():
            exams_in_week3 = [exam for exam in mods if exam in schedule and schedule[exam][0] in week3_days]
            if len(exams_in_week3) > 1:
                violations.append(f"❌ Module leader {leader} has more than one exam in week 3: {exams_in_week3}")
    
    ```
    
    This makes a list of all the days in the third weeks and then a list of exams that each module leader / examiner has on those days. If this is more than one exam it raises a violations
    
- Students who have more than 50% extra time shouldn't have more than one exam a day
    
    ```python
     # 6. Students with >50% extra time cannot have more than one exam on the same day
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
    
    ```
    
    This loops through all the students with 50% extra time and makes a dictionary of day and number of exams on this day. If a day has over one exam it raises a violation.
    
- Ideally students who have 25% extra time shouldn't have more than one exam a day (soft)
    
    ```python
        #7 soft Students with 25% extra time cannot have more than one exam on the same day
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
    ```
    
    This loops through all AEA students who don't have 50% or more extra time and does the same as above but only raises a soft error.
    
- No more than 2 exams on any slot (soft)
    
    ```python
            for exam in exams:
                day, slot,rooms = schedule[exam]
    
                if day <= 15:  # First two weeks
                    exam_in_slot[(day, slot)].append(exam)
    
            # Check for violations
            for date_slot, scheduled_exams in exam_in_slot.items():
                if len(scheduled_exams) >= 3:
                    violations.append(
                        f"⚠️ Soft warning: day/slot {date_slot} has {len(scheduled_exams)} exams scheduled: {scheduled_exams}"
                    )
    ```
    

# Room constraints - check

- Enough capacity to take everyone
    
    ```python
    # 1. Check room capacity sufficiency per exam
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
    ```
    
    This loop through every exam then obtains how many AEA and SEQ students are in it, it then loops through each room for that exams and sums total capacity. It then lists a violation if the AEA capacity or SEQ capacity is less than needed
    
- No room is being used in the same slot twice
    
    ```python
     room_schedule = defaultdict(list)  # key=(day, slot, room), value=list of exams
        for exam, (day, slot, rooms_) in exams_timetabled.items():
            for room in rooms_:
                room_schedule[(day, slot, room)].append(exam)
    
        for (day, slot, room), exams_in_room in room_schedule.items():
        		if room != "NON ME N/A": #skip the non me modules
    	        if len(exams_in_room) > 1:
                violations.append(
                    f"❌ Room '{room}' double-booked on day {day}, slot {slot} for exams: {exams_in_room}"
                )
    ```
    
    This makes a dictionary of each day slot and and room and if the there is more than one exam at the same day slot and in the same room it raises a violation
    
- PC exams must be in a PC room
    
    ```python
        # 3. Check computer-based exams are in computer rooms
        for exam, (day, slot, rooms) in exams_timetabled.items():
            if exam_types[exam] == "PC":  # Only check computer-based exams
                for room in rooms:
                    if "Computer" not in room_dict[room][0]:
                        violations.append(
                            f"❌ Computer-based exam '{exam}' assigned to non-computer room '{room}'"
                        )
    ```
    
    Loops through exams and checks wether it is a PC exam or not, if it is it checks the rooms and if they are not a computer room it raises a violation
    
- Check every exam has a room
    
    ```python
        # 4 Check every exam assigned at least one room
        for exam, (day, slot, rooms) in exams_timetabled.items():
            if not rooms:
                violations.append(f"❌ Exam '{exam}' has no assigned room!")
    
    ```
    
    This checks that for each exam in the schedule there is at least one room assigned to it.
    
- Non Pc exams not in PC rooms (Soft)
    
    ```python
        # 5 Check non PC exams are not in PC rooms
        for exam, (day, slot, rooms) in exams_timetabled.items():
            if exam_types[exam] != "PC":  # Only check non computer-based exams
                for room in rooms:
                    if "Computer" in room_dict[room][0]:
                        violations.append(
                            f"⚠️ Soft warning: '{exam}' assigned to computer room '{room}' and is not a computer exam"
                        )
    ```