# Exam timetabling - UROP

https://github.com/edwardb05/Timetabling-UROP - all the research and testing python scripts

https://github.com/edwardb05/streamlit_exe - The final executable

## Project Objectives:

- Develop a user-friendly Python script to generate ME3/4 exam timetables
- Enable timetable modifications following Aero Department consultation
- Support the addition of custom constraints

## [Inputs:](Handling%20data%2020c713b1f81d80529d8de3c3b51da11a.md)

- Excel file of students and their modules
- Excel file of module leaders
- Excel file of useful dates for the academic year

## Outputs:

- Excel spreadsheet with the dates of weeks 31-33 and when each exam is.

## [Mandatory constraints (Exams):](Constraints%2020b713b1f81d80d0bc1bf807a61b3e6b.md)

- No more than three exams in two days
- No more than four exams in five consecutive days (Monday to Friday)
- Core modules are fixed as students cannot have more than one exam on this day (In red on timetable)
- Other modules are fixed in date and time (In yellow on timetable )
- Module leaders can’t have more than one exam in the third week
- Students who have more than 50% extra time shouldn't have more than one exam a day

## [Mandatory constraints (Rooms):](Constraints%2020b713b1f81d80d0bc1bf807a61b3e6b.md)

- There must be capacity for all students taking exams and AEA students must have their own room
- A room cannot be used for more than one exam at the same time

## [Other (soft) constraints](Constraints%2020b713b1f81d80d0bc1bf807a61b3e6b.md):

- Ideally students who have 25% extra time shouldn't have more than one exam a day
- Spread out module leaders exams
- PC exams should be in computer rooms
- Minimize the amount of rooms used

## The software:

[https://lucid.app/lucidchart/57c3c56c-aa97-4974-bf91-29ce1d6b92a2/edit?beaconFlowId=4D9F2177E19CEB27&invitationId=inv_4c015f2d-b603-402a-92f8-e3688708b3fd&page=0_0#](https://lucid.app/lucidchart/57c3c56c-aa97-4974-bf91-29ce1d6b92a2/edit?beaconFlowId=4D9F2177E19CEB27&invitationId=inv_4c015f2d-b603-402a-92f8-e3688708b3fd&page=0_0#)

[Questions](Questions%20209713b1f81d8010a06bdda2455d1208.md)

[Research into methods](Research%20into%20methods%20209713b1f81d80219185c3f204e280b4.md)

[Experimenting with CSP](Experimenting%20with%20CSP%2020b713b1f81d800fae3be4bdc71357b1.md)

[Handling data](Handling%20data%2020c713b1f81d80529d8de3c3b51da11a.md)

[Constraints](Constraints%2020b713b1f81d80d0bc1bf807a61b3e6b.md)

[Displaying the timetable](Displaying%20the%20timetable%2020c713b1f81d80d3898be2e5cc8e7694.md)

[Checking algorithm](Checking%20algorithm%2020f713b1f81d80f18c35f973df0928a4.md)

[Streamlit app ](Streamlit%20app%20214713b1f81d80fa8f46e61f572fc681.md)

[Wrapping into an executable](Wrapping%20into%20an%20executable%2026f713b1f81d8094b756d9de4b23ecb1.md)

[Possible bugs](Possible%20bugs%20216713b1f81d808483a3d0d538ae557b.md)
