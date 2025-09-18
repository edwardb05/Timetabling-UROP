# Experimenting with CSP

Owner: Edward Brady

[https://github.com/edwardb05/Timetabling-UROP/blob/main/Testing CPS.ipynb](https://github.com/edwardb05/Timetabling-UROP/blob/main/Testing%20CPS.ipynb)

To experiment with CSP I chose to use CP-SAT solver from googles OR-tools package as it can solve for constraints and then also optimize to ensure module leaders have their exams spread out well.

## Using OR-tools packages:

```python
# Sample data

exams = ['Math', 'Physics', 'Chemistry', 'History','Geography','PE','Music']
students = {
    'Alice': ['Math', 'Physics','Geography','PE'],
    'Bob': ['Physics', 'Chemistry','Music'],
    'Charlie': ['Math', 'History','PE'],
}
days = ['Mon', 'Tue', 'Wed','Thursday']

# Fixed exam slots (exam: day index)
fixed_slots = {
    'History': 2  # History exam fixed on Wed
}
```

A simple tester to experiment, this data has some exams and three students. It states that history must be on a fixed day. 

```python

# Constraint 1: No student has more than 1 exam on the same day
for student, stu_exams in students.items():
    for i in range(len(stu_exams)):
        for j in range(i+1, len(stu_exams)):
            # exams for this student must be on different days
            model.Add(exam_schedule[stu_exams[i]] != exam_schedule[stu_exams[j]])
```

Adds a constraint to ensure no student has more than 1 exam on the same day.

```python
# Solve model
solver = cp_model.CpSolver()
status = solver.Solve(model)

if status == cp_model.FEASIBLE or status == cp_model.OPTIMAL:
    print("Exam schedule:")
    for exam in exams:
        assigned_day = solver.Value(exam_schedule[exam])
        print(f" - {exam}: {days[assigned_day]}")
else:
    print("No solution found.")

```

```python
Exam schedule:
 - Math: Mon
 - Physics: Wed
 - Chemistry: Tue
 - History: Wed
 - Geography: Tue
 - PE: Thursday
 - Music: Thursday
```

It solves it and then says which exams should be on which day.