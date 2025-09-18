# Constraints

Owner: Edward Brady

The Key part of using a CSP solver is making sure the constraints are set correctly, below is a list of all the constraints, the code used to set them and an explanation of what it does.

# Student Exam Constraints

## Students cannot have two exams at the same time

```python
for student, exams in student_exams.items():
#1 Loops through students
    for i in range(len(exams)):
        for j in range(i + 1, len(exams)):
            exam1 = exams[i]
            exam2 = exams[j]
            
            #2 Check whether the two exams are scheduled 
            same_day = exam_day[exam1] == exam_day[exam2]
            same_slot = exam_slot[exam1] == exam_slot[exam2]
            model.AddBoolOr([same_day.Not(), same_slot.Not()])
```

This filters through ‘core days’ which are days with a core exam on it and ensures that the only exam on this day is the core exam.

1. Loops through students and there exams
2. Ensure that the exams are either not in the same slot or not on the same day

## Core modules are fixed at students cannot have more than one exam on this day (In red on timetable)

```python
for student, exs in student_exams.items():#1 Loop students
		#2 Create a list of their core exams
    core_mods = [exam for exam in exs if exam in core_modules]
    other_mods = [exam for exam in exs if exam not in core_modules]
		#3 Loop through the core exams
    for exam in core_mods:
        for other in other_mods:
		        #4 Add constraint
            model.Add(exam_day[exam] != exam_day[other])
```

This filters through ‘core days’ which are days with a core exam on it and ensures that the only exam on this day is the core exam.

1. Loops through students
2. List their core modules
3. Loop through the exams that are core 
4. Ensure these are not on the same day as anything else

## Other modules are fixed in date and time (In yellow on timetable )

```python
for exam, fixed_day_slot in fixed_slots.items():
	model.Add(exam_day[exam] == fixed_day_slot[0])
	model.Add(exam_slot[exam] == fixed_day_slot[1])
	
	
```

Filters through each exam in the list of fixed ones and ensures they are on the correct days.

## Avoid scheduling an exam in slots with no exams such as bank holidays

```python

for exam in exams:
    for day, slot in no_exam_dates:
        model.AddForbiddenAssignments([exam_day[exam], exam_slot[exam]], [(day, slot)])
       
```

1. Loop through every exam
2. Loop through all the forbidden times
3. Forbid the assignment of an exam to those days

## No more than three exams in two days

```python
for student,ex in student_exams:
    for d in range(num_days - 1):  #1 Go through each pair of 2 days
        exams_in_2_days = []

        for exam in ex:  # Go through their exams
		        #2 Create the Boolean variables
            is_on_d = model.NewBoolVar(f'{student}_{exam}_on_day_{d}')
            is_on_d1 = model.NewBoolVar(f'{student}_{exam}_on_day_{d+1}')
            is_in_either = model.NewBoolVar(f'{student}_{exam}_on_day_{d}_or_{d+1}')

            #3 These say whether an exam is on day d or d+1
            model.Add(exam_day[exam] == d).OnlyEnforceIf(is_on_d)
            model.Add(exam_day[exam] != d).OnlyEnforceIf(is_on_d.Not())
            model.Add(exam_day[exam] == d+1).OnlyEnforceIf(is_on_d1)
            model.Add(exam_day[exam] != d+1).OnlyEnforceIf(is_on_d1.Not())

            #4 This says: if the exam is on either day, mark it as true
            model.AddBoolOr([is_on_d, is_on_d1]).OnlyEnforceIf(is_in_either)
            model.AddBoolAnd([is_on_d.Not(), is_on_d1.Not()]).OnlyEnforceIf(is_in_either.Not())

            exams_in_2_days.append(is_in_either)

        #5 Final rule: no more than 3 exams in this 2-day window
        model.Add(sum(exams_in_2_days) <= 3)
```

1. This applies a window over every two days in the exam period and ensures that each student does not have more than three exams in it. 
2. It does this by first creating three boolean variables to check if there is an exam on either the first or second day of the window and then check if it is in either.
3. We then go through each exam and check that the date of the exam is not one of the two days in the window. 
4. If the exam is in either of the day it is added to the list of exams in 2 days
5. Finally a constraint is added to ensure this list stays less than 3 items long.

## No more than four exams in five consecutive days (Monday to Friday)

```python
for student, exs in student_exams.items():
    for start_day in range(num_days - 4):#1 sliding window
        exams_in_window = []
        for exam in exs:#2 Each exam a student has
	        #3 Add in window variable
            in_window = model.NewBoolVar(f'{student}_{exam}_in_day_{start_day}_to_{start_day + 4}')

            model.AddLinearConstraint(exam_day[exam], start_day, start_day + 4).OnlyEnforceIf(in_window)

            before_window = model.NewBoolVar(f'{student}_{exam}_before_{start_day}')
            after_window = model.NewBoolVar(f'{student}_{exam}_after_{start_day + 4}')

            model.Add(exam_day[exam] < start_day).OnlyEnforceIf(before_window)
            model.Add(exam_day[exam] >= start_day).OnlyEnforceIf(before_window.Not())

            model.Add(exam_day[exam] > start_day + 4).OnlyEnforceIf(after_window)
            model.Add(exam_day[exam] <= start_day + 4).OnlyEnforceIf(after_window.Not())
						
						#4 Check whether exam in window
            model.AddBoolOr([before_window, after_window]).OnlyEnforceIf(in_window.Not())
						
						#5 Add exam to list
            exams_in_window.append(in_window)
				
				#6 Add constraint
        model.Add(sum(exams_in_window) <= 4)
```

This filters through every 5 day window for each student and ensures that they amount of exams they have in this window is less than 4.

1. Loop through every start day for the exam period
2. Looping through each exam that each student has
3. Initialize a Boolean variable for wether this exam is in the window or not
4. Check that the exams start day is not in the window
5. If the exam is in the window then add it to the list
6. This list cannot be greater than 4

## Module leaders can’t have more than one exam in the third week

```python
# 6. At most 1 exam in week 3 (days 13 to 20) per module leader
for leader, leader_exams in leader_courses.items(): #1 Loop through module leaders
    exams_in_week3 = []
    for exam in leader_exams: #2 Loop through each module leaders exams
        #3 Initialize variable
        in_week3 = model.NewBoolVar(f'{exam}_in_week3')

        model.AddLinearConstraint(exam_day[exam], 13, 20).OnlyEnforceIf(in_week3)
				
				#4 Check to see if exam in week 3
        before_week3 = model.NewBoolVar(f'{exam}_before_week3')
        after_week3 = model.NewBoolVar(f'{exam}_after_week3')

        model.Add(exam_day[exam] < 13).OnlyEnforceIf(before_week3)
        model.Add(exam_day[exam] >= 13).OnlyEnforceIf(before_week3.Not())

        model.Add(exam_day[exam] > 20).OnlyEnforceIf(after_week3)
        model.Add(exam_day[exam] <= 20).OnlyEnforceIf(after_week3.Not())

        model.AddBoolOr([before_week3, after_week3]).OnlyEnforceIf(in_week3.Not())
				#5 Add exam to list
        exams_in_week3.append(in_week3)
    
    #6 Add constraint
    model.Add(sum(exams_in_week3) <= 1)

```

1. Loop through a list of module leaders
2. Loop through each module leaders set of exams
3. Initialize a boolean variable to see if an exam is in week 3
4. Check wether the date of the exam is between the start and end days of week 3, pre-defined
5. If the exam is in the third week add it to the list
6. Add the constraint that there is only one or less exams in week 3 for each module leader

## Students who have more than 50% extra time shouldn't have more than one exam a day

```python
for student in extra_time_students_50:#1 Loop through extra time students
    for day in num_days:#2 Loop through days
        exams_on_day = []

        for exam in student_exams[student]:
		        #3 Initialize variable
            is_on_day = model.NewBoolVar(f'{student}_{exam}_on_day_{day}')
            #4 Check if exam is on day
            model.Add(exam_day[exam] == day).OnlyEnforceIf(is_on_day)
            model.Add(exam_day[exam] != day).OnlyEnforceIf(is_on_day.Not())
						#5 Add exam to list
            exams_on_day.append(is_on_day)

        #6 Add constraint 
        model.Add(sum(exams_on_day) <= 1)
```

1. Loop through a list of extra time students
2. Loop through every day of the exam season
3. Initialize a variable to see if the exam is on this day
4. Check if the exam is on this day
5. If it is add it to the list of exams this day
6. This list cannot exceed one

## Ideally students who have 25% extra time shouldn't have more than one exam a day (soft)

```python
 extra_time_25_penalties = []

for student in extra_time_students_25:#1 Loop through students
    for day in range(num_days): #2 Loop through days
        exams_on_day = []

        for exam in student_exams[student]:#3 Loop through exams
		        #4 Initialize variable
            is_on_day = model.NewBoolVar(f'{student}_{exam}_on_day_{day}')
            model.Add(exam_day[exam] == day).OnlyEnforceIf(is_on_day)
            model.Add(exam_day[exam] != day).OnlyEnforceIf(is_on_day.Not())
            exams_on_day.append(is_on_day)

        #5 Count exams
        num_exams = model.NewIntVar(0, len(exams_on_day), f'{student}_num_exams_day_{day}')
        model.Add(num_exams == sum(exams_on_day))

        # 6 Calculate if >1 exam
        has_multiple_exams = model.NewBoolVar(f'{student}_more_than_one_exam_day_{day}')
        model.Add(num_exams >= 2).OnlyEnforceIf(has_multiple_exams)
        model.Add(num_exams < 2).OnlyEnforceIf(has_multiple_exams.Not())

        #7 Assign a peanlty if they have more than 1 exam
        penalty = model.NewIntVar(0, 1, f'{student}_penalty_day_{day}')
        model.Add(penalty == 1).OnlyEnforceIf(has_multiple_exams)
        model.Add(penalty == 0).OnlyEnforceIf(has_multiple_exams.Not())

        extra_time_25_penalties.append(penalty)
```

1. Loop through a list of extra time students with 25%
2. Loop through every day of the exam season
3. Loop through all of the students exams
4. Initialize the variable for whether an exam is on this day and check if its true
5. Count the exams on the day 
6. Calculate if there is more than 1 exam on a day
7. Assign a penalty if the exam is greater than 1

## Spread out module leaders exams (soft)

```python
spread_penalties =[]

for leader in module_leaders:#1 Loop through module leaders
    mods = leader_modules[leader]
    
    for i in range(len(mods)):#2 loop through modules
        for j in range(i+1, len(mods)):
            m1 = mods[i]
            m2 = mods[j]

            #3 Calculate absolute day difference
            diff = model.NewIntVar(-21, 21, f'{m1}_{m2}_diff')
            abs_diff = model.NewIntVar(0, 21, f'{m1}_{m2}_abs_diff')
						#4 Add difference to model
            model.Add(diff == exam_day[m1] - exam_day[m2])
            model.AddAbsEquality(abs_diff, diff)

            #5 Create penalty variable
            close_penalty = model.NewIntVar(0, 5, f'{m1}_{m2}_penalty')

            #6 Create Boolean conditions
            is_gap_3 = model.NewBoolVar(f'{m1}_{m2}_gap3')
            is_gap_2 = model.NewBoolVar(f'{m1}_{m2}_gap2')
            is_gap_1 = model.NewBoolVar(f'{m1}_{m2}_gap1')
            is_gap_0 = model.NewBoolVar(f'{m1}_{m2}_gap0')
						#7 Set the true condition
            model.Add(abs_diff == 3).OnlyEnforceIf(is_gap_3)
            model.Add(abs_diff != 3).OnlyEnforceIf(is_gap_3.Not())

            model.Add(abs_diff == 2).OnlyEnforceIf(is_gap_2)
            model.Add(abs_diff != 2).OnlyEnforceIf(is_gap_2.Not())

            model.Add(abs_diff == 1).OnlyEnforceIf(is_gap_1)
            model.Add(abs_diff != 1).OnlyEnforceIf(is_gap_1.Not())

            model.Add(abs_diff == 0).OnlyEnforceIf(is_gap_0)
            model.Add(abs_diff != 0).OnlyEnforceIf(is_gap_0.Not())

            #8 Assign penalty values based gap
            model.Add(close_penalty == 1).OnlyEnforceIf(is_gap_3)
            model.Add(close_penalty == 3).OnlyEnforceIf(is_gap_2)
            model.Add(close_penalty == 4).OnlyEnforceIf(is_gap_1)
            model.Add(close_penalty == 5).OnlyEnforceIf(is_gap_0)

            #9 no penalty if gap ≥ 4 and not equal to 0–3
            model.Add(close_penalty == 0).OnlyEnforceIf(
                is_gap_3.Not(), is_gap_2.Not(), is_gap_1.Not(), is_gap_0.Not()
            )
						#10 Add penalty to total penalties
            spread_penalties.append(close_penalty)

```

1. Loop through the module leaders
2. Loop through their respective modules
3. Calculate the difference in days between all of their modules
4. Assign the absolute value of this to the module
5. Create the penalty variable
6. Create the boolean variable to check for different gap lengths.
7. Check to see how long the gap is and set this boolean variable to true
8. Assign a penalty value based on gap length
9. Theres no penalty if the gap is longer than 4
10. Add this gaps penalty to the total spread penalties

## Have no exams on Tuesday and Wednesday morning in week 3 (soft)

```python
#Soft constraint to ensure no exams on some days
#1 Initiated list of penalties
soft_day_penalties = []

for exam in exams:#2 Loop exams
    for day, slot in no_exam_dates_soft:
        is_on_soft_day = model.NewBoolVar(f'{exam}_on_soft_day_{day}_{slot}')

        #3 Boolean variables for day and slot matches
        day_match = model.NewBoolVar(f'{exam}_day_eq_{day}')
        slot_match = model.NewBoolVar(f'{exam}_slot_eq_{slot}')
        
        model.Add(exam_day[exam] == day).OnlyEnforceIf(day_match)
        model.Add(exam_day[exam] != day).OnlyEnforceIf(day_match.Not())

        model.Add(exam_slot[exam] == slot).OnlyEnforceIf(slot_match)
        model.Add(exam_slot[exam] != slot).OnlyEnforceIf(slot_match.Not())

        #4 is_on_soft_day = day_match AND slot_match
        model.AddBoolAnd([day_match, slot_match]).OnlyEnforceIf(is_on_soft_day)
        model.AddBoolOr([day_match.Not(), slot_match.Not()]).OnlyEnforceIf(is_on_soft_day.Not())

        #5 Penalty: 10 if scheduled on soft day
        penalty = model.NewIntVar(0, 10, f'{exam}_penalty_soft_day_{day}_{slot}')
        model.Add(penalty == 10).OnlyEnforceIf(is_on_soft_day)
        model.Add(penalty == 0).OnlyEnforceIf(is_on_soft_day.Not())

        soft_day_penalties.append(penalty)
```

1. Initiate the list of penalties
2. Loop through all the exams
3. Create the boolean conditions to check if an exam is on a soft day or soft slot
4. Check if its on both
5. Add a penalty if it is 

## Amount of exams in one slot preferably one or 2, in first two weeks (soft)

```python
soft_slot_penalties = []

for day in range(15):  #1 First two weeks only
    for slot in slots:  
        exams_in_slot = []
			
				# 2 Make a list of all exams in a slot
        for exam in exams:
            is_scheduled_day = model.NewBoolVar(f'{exam}_is_on_day{day}')
            is_scheduled_slot = model.NewBoolVar(f'{exam}_is_on_slot{slot}')
            
            model.Add(exam_day[exam] == day).OnlyEnforceIf(is_scheduled_day)
            model.Add(exam_day[exam] != day).OnlyEnforceIf(is_scheduled_day.Not())
            
            model.Add(exam_slot[exam] == slot).OnlyEnforceIf(is_scheduled_slot)
            model.Add(exam_slot[exam] != slot).OnlyEnforceIf(is_scheduled_slot.Not())
            
            is_scheduled_here = model.NewBoolVar(f'{exam}_on_day{day}_slot{slot}')
            model.AddBoolAnd([is_scheduled_day, is_scheduled_slot]).OnlyEnforceIf(is_scheduled_here)
            model.AddBoolOr([is_scheduled_day.Not(), is_scheduled_slot.Not()]).OnlyEnforceIf(is_scheduled_here.Not())
            exams_in_slot.append(is_scheduled_here)

        # 3 Count number of exams scheduled in this (day, slot)
        num_exams_here = model.NewIntVar(0, len(exams), f'count_day{day}_slot{slot}')
        model.Add(num_exams_here == sum(exams_in_slot))

        # 4 Calculate penalties
        is_three = model.NewBoolVar(f'is_three_day{day}_slot{slot}')
        is_four_or_more = model.NewBoolVar(f'is_four_plus_day{day}_slot{slot}')

        model.Add(num_exams_here == 3).OnlyEnforceIf(is_three)
        model.Add(num_exams_here != 3).OnlyEnforceIf(is_three.Not())

        model.Add(num_exams_here >= 4).OnlyEnforceIf(is_four_or_more)
        model.Add(num_exams_here < 4).OnlyEnforceIf(is_four_or_more.Not())

        #5 Apply penalties
        penalty_three = model.NewIntVar(0, 5, f'penalty_three_day{day}_slot{slot}')
        penalty_four = model.NewIntVar(0, 100, f'penalty_four_day{day}_slot{slot}')

        model.Add(penalty_three == 5).OnlyEnforceIf(is_three)
        model.Add(penalty_three == 0).OnlyEnforceIf(is_three.Not())

        model.Add(penalty_four == 100).OnlyEnforceIf(is_four_or_more)
        model.Add(penalty_four == 0).OnlyEnforceIf(is_four_or_more.Not())

        soft_slot_penalties.append(penalty_three)
        soft_slot_penalties.append(penalty_four)

```

1. Loop through each day and slot in the first two weeks (day 15)
2. Loop through all the exams and check whether they are in this slot or not
3. Count the length of the list and thus the number of exams on this exact slot
4. Define the variables and calculate which one to use
5. Apply penalty values

## How to minimize soft constraints

```python
model.Minimize(sum(spread_penalties + soft_day_penalties*soft_day_penalty+   extra_time_25_penalties*extra_time_penalty+room_surplus+ soft_slot_penalties*soft_slot_penalty+ non_pc_exam_penalty))
```

This adds the  lists of penalties, multiplies the necessary ones by their multipliers and ensures they are minimized 

# Room constraints - check

## No rooms for Non Mech Eng modules in yellow on the timetable

```python
#1 Loop exams
for exam in exams:
    if exam in Fixed_modules and exam not in Core_modules: #2 check if exam is non mech eng
        model.Add(exam_room[(exam, 'N/A')] ==1)  #3 Assign to N/A room if fixed module
    else:
          model.Add(exam_room[(exam, 'N/A')] == 0)  #4 Do not assign to N/A room if not fixed module

```

1. Loop through every exam
2. Check if an exam is mech Eng by seeing if it is in the fixed modules and is not one of the core modules
3. Set the hard constraint that if non Mech Eng it gets assigned the N/A room
4. Set the hard constraint that Mech Eng exams cannot have a N/A 

## Enough capacity to take everyone

```python
for exam in exams: #1 Loop through
        #2 Calculate capacity's for each room
        AEA_capacity = sum(
            rooms[room][1] * exam_room[(exam, room)]
            for room in rooms if "AEA" in rooms[room][0]
        )

        SEQ_capacity = sum(
            rooms[room][1] * exam_room[(exam, room)]
            for room in rooms if "SEQ" in rooms[room][0]
        )

				
        AEA_students = exam_counts[exam][0]
        SEQ_students = exam_counts[exam][1]

        #3 Add Constraint
        model.Add(AEA_capacity >= AEA_students)
        model.Add(SEQ_capacity >= SEQ_students)

```

1. Loop through every exam
2. The capacity for AEA students and other students is calculated, the amount of students taking the exam is also calculated this is slightly difficult as exam_room is a boolean variable so the capcity of the room is multiplied by 1 if it is used and 0 if it isn’t
3. Set the hard constraint that Capacity has to be higher than the amount of students

## No room is being used in the same slot twice

```python
for d in range(num_days):
    for s in range(num_slots):#1Loop through each day and slot
        for room in rooms: 
            exams_in_room_time = []
            for exam in exams:#2Loop through exams and rooms
 

                # 3Create bool var
                exam_at_day = model.NewBoolVar(f'{exam}_on_day_{d}')
                model.Add(exam_day[exam] == d).OnlyEnforceIf(exam_at_day)
                model.Add(exam_day[exam] != d).OnlyEnforceIf(exam_at_day.Not())

                exam_at_slot = model.NewBoolVar(f'{exam}_on_slot_{s}')
                model.Add(exam_slot[exam] == s).OnlyEnforceIf(exam_at_slot)
                model.Add(exam_slot[exam] != s).OnlyEnforceIf(exam_at_slot.Not())
								#4 Create variable for exams at time
                exam_at_time = model.NewBoolVar(f'{exam}_on_{d}_{s}')
                model.AddBoolAnd([exam_at_day, exam_at_slot]).OnlyEnforceIf(exam_at_time)
                model.AddBoolOr([exam_at_day.Not(), exam_at_slot.Not()]).OnlyEnforceIf(exam_at_time.Not())

                # 5 Create variable for room
                assigned_and_scheduled = model.NewBoolVar(f'{exam}_in_{room}_at_{d}_{s}')
                model.AddBoolAnd([exam_room[(exam, room)], exam_at_time]).OnlyEnforceIf(assigned_and_scheduled)
                model.AddBoolOr([exam_room[(exam, room)].Not(), exam_at_time.Not()]).OnlyEnforceIf(assigned_and_scheduled.Not())
								#6 add list of exams in room at time
                exams_in_room_time.append(assigned_and_scheduled)

            # 7 add constraint
            model.AddAtMostOne(exams_in_room_time)

```

1. Loop through all days and time slots
2. Loop through all the exams and rooms
3. Create a variable to see if the exam is at this day or slot
4. Create a variable to see if the exam is at both this day and slot
5. Create a variable to see if the exam is scheduled in this room at this time 
6. If the rooms is at this time and in this room add it to the list of exams in this room and time
7. Add a constraint to ensure this list is only one exam long

## PC exams must be in a PC room

```python
for exam in exams: #1 Loop exams
    if exam_types[exam] == "PC": #Check exam
        for room in rooms:
            uses = rooms[room][0]  

            if "Computer" not in uses:#3 Check if room can be used as a computer room
                #4 add constraint
                model.Add(exam_room[(exam, room)] == 0)
```

1. Loop through exams
2. Check if the exam is a PC Exam
3. Loop through the rooms and see if they are a computer room
4. If they are not add constraint that forbids it being used for this exam

## Minimise amount of rooms used for each exam and ensure each exam has at least one room

```python

room_surplus = [] #1 Initialize list of surplus
for exam in exams:#2 Loop through exams
		#3 Add constraint the each exam has more than 1 room
    model.Add(sum(exam_room[(exam, room)] for room in rooms) >= 1)
    
    #4Create integer for amount of rooms
    rooms_len = model.NewIntVar(0, 9, f'rooms for {exam}')
    
    model.Add(rooms_len == sum(exam_room[(exam, room)]for room in rooms))

     #5 Create penalty variable
    rooms_penalty = model.NewIntVar(0, 15, f'{exam}_room_surplus_penalty')

     #6 Create Boolean conditions
    is_room_length_greater_6 = model.NewBoolVar(f'{exam}_has_six_or_more_rooms')
    is_room_length_5 = model.NewBoolVar(f'{exam}_has_five_rooms')
    is_room_length_4 = model.NewBoolVar(f'{exam}_has_four_rooms')
    is_room_length_3 = model.NewBoolVar(f'{exam}_has_three_rooms')
    
		#7 Set the true condition
		model.Add(rooms_len >= 6).OnlyEnforceIf(is_room_length_greater_6)
    model.Add(rooms_len <= 5).OnlyEnforceIf(is_room_length_greater_6.Not())
		
    model.Add(rooms_len == 5).OnlyEnforceIf(is_room_length_5)
    model.Add(rooms_len != 5).OnlyEnforceIf(is_room_length_5.Not())

    model.Add(rooms_len == 4).OnlyEnforceIf(is_room_length_4)
    model.Add(rooms_len != 4).OnlyEnforceIf(is_room_length_4.Not())
        
    model.Add(rooms_len == 3).OnlyEnforceIf(is_room_length_3)
    model.Add(rooms_len != 3).OnlyEnforceIf(is_room_length_3.Not())

     #8 Assign penalty values based gap
    model.add(rooms_penalty == 15).OnlyEnforceIf(is_room_length_greater_6)
    model.Add(rooms_penalty == 9).OnlyEnforceIf(is_room_length_5)
    model.Add(rooms_penalty == 6).OnlyEnforceIf(is_room_length_4)
    model.Add(rooms_penalty == 4).OnlyEnforceIf(is_room_length_3)

     #9 no penalty if less than 3 rooms 
    model.Add(rooms_penalty == 0).OnlyEnforceIf(
                is_room_length_3.Not(), is_room_length_4.Not(), is_room_length_5.Not(), is_room_length_greater_6.Not(),
            )
						#10 Add penalty to total penalties
    room_surplus.append(rooms_penalty)
```

1. Initialise the list of penalties
2. Loop through each exams
3. Add a hard constraint to ensure each exam has at least one room
4. Create the integer for the amount of rooms an exam has
5. Create the penalty variable with max penalty 15
6. Create the 4 boolean conditions we will be testing for
7. Add the logic for all the conditions 
8. Assign a penalty based on the amount of rooms
9. Add no penalty if there's 2 or less rooms
10. Add the penalties to the total penalties

## Penalise using computer rooms for non pc exam

```python

non_pc_exam_penalty = []

#1 Find computer rooms
computer_rooms = [room for room in rooms if "Computer" in rooms[room][0]]

#2 Loop through exams
for exam in exams:
		#3 if not a PC exams
    if exam_types[exam] != "PC":
		    #4 Check each computer room
        for room in computer_rooms:
            #5 Create a Boolean 
            penalty_var = model.NewBoolVar(f"non_pc_exam_in_pc_room_{exam}_{room}")
            
            #6 Assign penalty 
            model.Add(exam_room[(exam, room)] == 1).OnlyEnforceIf(penalty_var)
            model.Add(exam_room[(exam, room)] != 1).OnlyEnforceIf(penalty_var.Not())

            #7 Add penality
            non_pc_exam_penalty.append(5 * penalty_var)
```

1. Initialise the list of penalties
2. Loop through each exams
3. Check to see if the exam is not a pc exam
4. Check to see if any of the computer rooms are being used 
5. Create a boolean penalty variable
6. Assign true if a computer room is being used for a new computer exam
7. Add the penalty to the list of penalties.