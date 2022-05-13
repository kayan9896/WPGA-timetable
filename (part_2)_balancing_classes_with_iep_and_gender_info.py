# -*- coding: utf-8 -*-
"""
## Timetabling Program for West Point Grey Academy
"""

# Import Python Modules

import time
import numpy as np
import pandas as pd
from random import random
from random import shuffle
from ortools.linear_solver import pywraplp

# Import the Input File with the 2022-2023 Student and Course Data.  

InputFile = pd.ExcelFile("WPGA 2022-2023 Data.xlsx") 
InputMatrix = pd.read_excel(InputFile, "Data") 
InputInfo = InputMatrix.values.tolist()

# Generate the list of courses and list of students.  Sort both lists.
# Let n be the number of students and m be the number of courses.

CourseList = []
for i in range(len(InputInfo)):
    CourseName = InputInfo[i][1]
    if not pd.isna(CourseName):
        if not CourseName in CourseList:
            CourseList.append(CourseName)
CourseList.sort()

StudentList = []
for i in range(len(InputInfo)):
    StudentID = InputInfo[i][31]
    if not StudentID in StudentList:
        StudentList.append(StudentID)
StudentList.sort()

n = len(StudentList)
m = len(CourseList)

print("There are", n, "students to be timetabled into", m, "courses")

# For each course j, let CourseSections[j] be the number of sections of that course.
# For example, CourseSections[15] = 2 since Course #15 (AP Psychology) has 2 sections.

CourseSections = [0 for j in range(m)]
for j in range(m):
    CourseName = InputInfo[j][1]
    NumberOfSections = int(InputInfo[j][5])
    CourseIndex = CourseList.index(CourseName)
    CourseSections[CourseIndex] = NumberOfSections
    

# For each student i, let StudentChoices[i] be the list of courses chosen by that student.

StudentChoices = [ [] for i in range(n)]
for i in range(len(InputInfo)):
    StudentID = InputInfo[i][31]
    StudentIndex = StudentList.index(StudentID)
    CourseName = InputInfo[i][35]
    CourseIndex = CourseList.index(CourseName)
    StudentChoices[StudentIndex].append(CourseIndex)

# Generate the list of forbidden (Course,Block) assignments.  If [j,k] appears in the
# ForbiddenAssignments list, then no section s of course j can be offered in block k.

ForbiddenAssignments = []

for x in range(m):
    Requirement = InputInfo[x][8]
    CourseName = InputInfo[x][1]
    if not pd.isnull(Requirement):
        j = CourseList.index(CourseName)
        if Requirement == '1A/2A':
            for k in [2,3,4,6,7,8,9]: ForbiddenAssignments.append([j,k])      
        elif Requirement == '1A/1B':
            for k in [3,4,5,6,7,8,9]: ForbiddenAssignments.append([j,k])
        elif Requirement == '1A/1B/2A/2B':
            for k in [3,4,7,8,9]: ForbiddenAssignments.append([j,k])
        elif Requirement == '2C/2D/2E':
            for k in [1,2,3,4,5,6]: ForbiddenAssignments.append([j,k])
        
        
# Generate the list of required (Section,Course,Block) assignments.  If [s,j,k] appears in the
# RequiredAssignments list, then section s of course j must be offered in block k.

RequiredAssignments = []

for x in range(m):
    Requirement = InputInfo[x][8]
    CourseName = InputInfo[x][1]
    if not pd.isnull(Requirement):
        j = CourseList.index(CourseName)
        if Requirement == '1A': RequiredAssignments.append([1,j,1])
        elif Requirement == '1B': RequiredAssignments.append([1,j,2])
        elif Requirement == '1C': RequiredAssignments.append([1,j,3])
        elif Requirement == '1D': RequiredAssignments.append([1,j,4])
        elif Requirement == '2A': RequiredAssignments.append([1,j,5])
        elif Requirement == '2B': RequiredAssignments.append([1,j,6])

j= CourseList.index("Study Block")
for s in [1,2,3,4,5,6,7,8,9]:
    RequiredAssignments.append([s,j,s])
    
j = CourseList.index("Study Block2")
for s in [1,2,3,4,5,6,7,8,9]:
    RequiredAssignments.append([s,j,s])
    
j = CourseList.index("Supervised Support Block")
for s in [1,2,3,4,5,6,7,8,9]:
    RequiredAssignments.append([s,j,s])

# For each course j, let CourseRequestList[j] be the list of students who requested that course.
# Also, let CourseRequestTotal[j] be the total number of students who requested that course.

CourseRequestList = [ [] for j in range(m)]
CourseRequestTotal = [ 0 for j in range(m)]
for i in range(n):
    for j in StudentChoices[i]:
        CourseRequestTotal[j] += 1
        CourseRequestList[j].append(i) 
        

# For each course j, let CoursePriority[j] be the priority level of that course.
# In the absence of any other information, assume CoursePriority[j] = 0 for all courses j.

CoursePriority = [0 for j in range(m)]


# For each course j, let PossibleTeachers[j] be the list of possible teachers for that course.
# Leave any commas and slash marks as they are.

PossibleTeachers = [ [] for j in range(m)]
for j in range(m):
    CourseName = InputInfo[j][1]
    CourseIndex = CourseList.index(CourseName)
    TeacherInfo = InputInfo[j][6]
    if pd.notna(TeacherInfo):
        PossibleTeachers[CourseIndex] = TeacherInfo
        

# From Column N of the Input Excel file, generate the list of teachers and sort this list.

TeacherList = []
for i in range(len(InputInfo)):
    TeacherName = InputInfo[i][13]
    if pd.notna(TeacherName):
        TeacherList.append(TeacherName)
TeacherList.sort()

    
# For each teacher t, let TeacherCourses[t] be the list of courses that MUST be taught
# by that teacher.  Each course is a number in range(m), based on the index of the course name
# in the variable CourseList.  For all courses with multiple options separated by slash marks / 
# (e.g. McAllister/JohnsonCalvert), leave those blank.  Only include courses where there is 
# one assigned teacher (e.g. Harms) or multiple teachers separated by commas (Bendl, Boland)

TeacherCourses = [ [] for t in range(len(TeacherList))]
for j in range(m):
    CourseName = InputInfo[j][1]
    CourseIndex = CourseList.index(CourseName)
    TeacherInfo = InputInfo[j][6]
    if pd.notna(TeacherInfo):
        TeacherSplit = TeacherInfo.split('/')
        if len(TeacherSplit) == 1:
            if TeacherSplit not in [["ANY"], ["PESTAFF"], ["Do not schedule"]]:
                TeacherName = TeacherSplit[0]
                if ',' in TeacherName:
                    AllTeachers = TeacherName.split(', ')
                    for Teacher in AllTeachers:
                        t = TeacherList.index(Teacher)
                        TeacherCourses[t].append(CourseIndex)
                else:
                    t = TeacherList.index(TeacherName)
                    TeacherCourses[t].append(CourseIndex)
                    
                    
# For each grade g in [8,9,10,11,12], define StudentsPerGrade[g] to be the list of students 
# in that grade based on the index of the student name in the variable StudentList.
 
StudentsPerGrade = [ [] for i in range(13)]
for i in range(len(InputInfo)):
    StudentID = InputInfo[i][31]
    StudentIndex = StudentList.index(StudentID)
    StudentGrade = InputInfo[i][34]
    if StudentGrade == 'Grade 8': StudentsPerGrade[8].append(StudentIndex)
    if StudentGrade == 'Grade 9': StudentsPerGrade[9].append(StudentIndex)
    if StudentGrade == 'Grade 10': StudentsPerGrade[10].append(StudentIndex)
    if StudentGrade == 'Grade 11': StudentsPerGrade[11].append(StudentIndex)
    if StudentGrade == 'Grade 12': StudentsPerGrade[12].append(StudentIndex)
StudentsPerGrade[8] = list(set(StudentsPerGrade[8]))
StudentsPerGrade[9] = list(set(StudentsPerGrade[9]))
StudentsPerGrade[10] = list(set(StudentsPerGrade[10]))
StudentsPerGrade[11] = list(set(StudentsPerGrade[11]))
StudentsPerGrade[12] = list(set(StudentsPerGrade[12]))


# For each course j, let RoomLimit[j] be the maximum capacity of that course
# If a course j has Room Requirement == "General", assume that RoomLimit[j] = 22.
# Otherwise, take the MAXIMUM value of the options.  

RoomLimit = [0 for j in range(m)]
RoomChoices = ['' for j in range(m)]

ClassroomList = [['General', 22]]

for k in range(len(InputInfo)):
    Room = InputInfo[k][18]
    Cap = InputInfo[k][20]
    if not pd.isnull(Room):
        ClassroomList.append([str(Room), int(Cap)])
        
for j in range(m):
    CourseName = InputInfo[j][1]
    if not pd.isnull(CourseName):
        CourseIndex = CourseList.index(CourseName)
        RoomOptions = str(InputInfo[j][9])
        RoomChoices[CourseIndex] = RoomOptions.split('/')
        flag=0
        for k in range(len(ClassroomList)):
            if ClassroomList[k][0]==RoomOptions:
                RoomLimit[CourseIndex] = ClassroomList[k][1]
                flag=1
        if flag==0:
            if ',' in RoomOptions:
                RoomLimit[CourseIndex] = 24
            else:
                PossibleRooms = RoomOptions.split('/')            
                for RoomNum in PossibleRooms:
                    for k in range(len(ClassroomList)):
                        if ClassroomList[k][0]==RoomNum:
                            if ClassroomList[k][1]>RoomLimit[CourseIndex]:
                                RoomLimit[CourseIndex] = ClassroomList[k][1]


# Three exceptions: set RoomLimit[j] = 100 for the following three courses:
# "Varsity Sport PE 10-12", "Study Block", and "Study Block2".

for CourseName in ['Varsity Sport PE 10-12', 'Study Block', 'Study Block2']:
    CourseIndex = CourseList.index(CourseName)
    RoomLimit[CourseIndex] = 100

    

# Determine the set of courses belonging to each of the five departments below

Departments = ["English", "Mathematics", "Languages", "Science", "Social Studies"]
DepartmentCourses = [[] for d in range(5)]

for j in range(m):
    CourseName = InputInfo[j][1]
    CourseIndex = CourseList.index(CourseName)
    DepartmentName = InputInfo[j][0] 
    for d in range(5):
        if DepartmentName == Departments[d]:
            DepartmentCourses[d].append(CourseIndex)
            
            
# Think of Geology as a Mathematics course rather than a Science course as Geology
# takes place in a math classroom.

DepartmentCourses[1].append(CourseList.index("Geology 12"))
DepartmentCourses[3].remove(CourseList.index("Geology 12"))

# Manual changes that need to be made to the data, based on the information
# provided in Ralph's Excel sheet.

j = CourseList.index("Materials Design 8.")
RoomLimit[j] = 15

j = CourseList.index("Visual Arts 9.")
RoomLimit[j] = 23

j = CourseList.index("Science 10x")
RoomLimit[j] = 25

j = CourseList.index("Theatre Company 10, 11, 12")
t = TeacherList.index("Penner-Tovey")
RoomLimit[j] = 50
TeacherCourses[t].append(j)

j = CourseList.index("Varsity Sport PE 10-12")
CourseSections[j] = 1
t = TeacherList.index("McCauley")
TeacherCourses[t].append(j)
t = TeacherList.index("GaringerD")
TeacherCourses[t].append(j)

j = CourseList.index("Global Studies 11/12 Seminar")
t = TeacherList.index("Liu")
TeacherCourses[t].append(j)
t = TeacherList.index("Johnston")
TeacherCourses[t].append(j)

j = CourseList.index("Environmental Science 12")
t = TeacherList.index("Harding")
TeacherCourses[t].append(j)

t = TeacherList.index("Liu")
for j in TeacherCourses[t]:
    for k in [5,6]: ForbiddenAssignments.append([j,k])
        
t = TeacherList.index("Logher")
for j in TeacherCourses[t]:
    for k in [1,2,5,6,7,8,9]: ForbiddenAssignments.append([j,k])
        
t = TeacherList.index("Penner-Tovey")
for j in TeacherCourses[t]:
    for k in [5,6,7,8,9]: ForbiddenAssignments.append([j,k])
        
t = TeacherList.index("McCauley")
for j in TeacherCourses[t]:
    for k in [6]: ForbiddenAssignments.append([j,k])

t = TeacherList.index("Elmer")
for j in TeacherCourses[t]:
    for k in [3,4,7,8,9]: ForbiddenAssignments.append([j,k])
        
t = TeacherList.index("Point")
for j in TeacherCourses[t]:
    for k in [3,4,7,8,9]: ForbiddenAssignments.append([j,k])
        
t = TeacherList.index("Goddard")
for j in TeacherCourses[t]:
    for k in [2,3,4,6,7,8,9]: ForbiddenAssignments.append([j,k])
        
# NOTE: we might fix this constraint later, if we can get a better result by switching
# Pope's required teaching blocks
t = TeacherList.index("Pope")
for j in TeacherCourses[t]:
    for k in [5,6,7,8,9]: ForbiddenAssignments.append([j,k])
        
        
j = CourseList.index("Physical and Health Education 8.")
CourseSections[j] = 1
RoomLimit[j] = 80
RequiredAssignments.append([1,j,2])
    
j = CourseList.index("Physical and Health Education 9.")
CourseSections[j] = 1
RoomLimit[j] = 80
RequiredAssignments.append([1,j,1])

j = CourseList.index("Physical and Health Education 10")
CourseSections[j] = 1
RoomLimit[j] = 80
RequiredAssignments.append([1,j,6])

j = CourseList.index("Active Living 11/12")
for k in [1,2,5,6]:
    ForbiddenAssignments.append([j,k])

j = CourseList.index("Active Living 11/12 - Individual Pursuits")
for k in [1,2,3,4,5,6]:  
    ForbiddenAssignments.append([j,k])

StudentMatrix = pd.read_excel(InputFile, usecols='AF:AL')
StudentInfo = StudentMatrix.values.tolist()

CourseMatrix = pd.read_excel(InputFile, usecols='A:U')
CourseMatrix = CourseMatrix[CourseMatrix['Department'].notnull()]
CourseInfo = CourseMatrix.values.tolist()


# GenderInfo[15] = 1 means student15 is male, and GenderInfo[20] = 0 means student20 is female
# IEP[i, j] will be a binary indicator about whether student i for course j is IEP,
# so if IEP[0][5] = 1, it means student0 has IEP for course5.
GenderInfo = [0 for _ in range(n)]
IEP = [[0 for _ in range(m)] for _ in range(n)]

for i in range(len(StudentMatrix)):
    row = StudentMatrix.iloc[i]
    StudentID = row["Hoshino Student ID"]
    StudentIndex = StudentList.index(StudentID)

    individualGender = row["Gender"].strip()
    if individualGender == "Male":
        GenderInfo[StudentIndex] = 1
    elif individualGender == "Female":
        GenderInfo[StudentIndex] = 0
    else:
        print("ERROR! Unknown gender")

    individualIEP = row["IEP Flag"]
    CourseName = row["Title Translation"]
    CourseIndex = CourseList.index(CourseName)
    if pd.notna(individualIEP):
        IEP[StudentIndex][CourseIndex] = 1

# Create our preference matrix P[i,j], where i is in range(n) and j is in range(m).
# P[i,j] is the preference for student i taking course j.

P = np.zeros((n,m), dtype=int)


# Assume the following weights for the "elective" courses: +2 for Gr.8, +5 for Gr.9, 
# +12 for Gr.10, +25 for Gr.11, +40 for Gr. 12.

for i in StudentsPerGrade[8]:
    for j in StudentChoices[i]: P[i,j] = 20
for i in StudentsPerGrade[9]:
    for j in StudentChoices[i]: P[i,j] = 30
for i in StudentsPerGrade[10]:
    for j in StudentChoices[i]: P[i,j] = 40
for i in StudentsPerGrade[11]:
    for j in StudentChoices[i]: P[i,j] = 50
for i in StudentsPerGrade[12]:
    for j in StudentChoices[i]: P[i,j] = 60
        
    
# Manually fix the preference coefficients for the unlucky students who did not get into all of their courses

UnluckyStudents = [169, 135, 202, 139, 183, 248, 250, 246, 377, 83, 344, 249, 123, 272, 19, 227, 361, 362, 34, 86, 401, 35, 297, 163, 339, 21, 398]
for i in UnluckyStudents:
    for k in range(len(StudentChoices[i])):
        j = StudentChoices[i][k]
        P[i,j] = 10 - k
        
        
# Overwrite the above preference coefficients for the following cases:
# Study blocks count as +1
# All cancelled courses (e.g. Beginner Spanish) and out-of-the-timetable courses 
# (e.g. Choral Music) are given weight 0 since these 0-section courses are not in the timetable.

for j in range(m):
    if CourseList[j] in ['Study Block', 'Study Block2']:
        for i in range(n):
            if P[i,j]>0: P[i,j] = 1
    if CourseSections[j] == 0:
        for i in range(n):
            if P[i,j]>0: P[i,j] = 0

# Decide which courses will need to have IEPs considered.  Ralph's rule is
# only courses where 15% of students (or more) have IEPs.

IEPcourses = []
for j in range(m):
    IEPcount = 0
    for i in range(n):
        if IEP[i][j] == 1:
            IEPcount +=1
    if CourseRequestTotal[j]>0:
        if IEPcount/CourseRequestTotal[j]>=0.15:
            IEPcourses.append(j)

# Create Hill-Climbing Program

def HillClimber(XSet, FixedNumber): 
    
    solver = pywraplp.Solver('Final Project', pywraplp.Solver.CBC_MIXED_INTEGER_PROGRAMMING)
    
    Students = range(len(StudentList))
    Courses = range(len(CourseList))
    Teachers = range(len(TeacherList))
    
    Sections = [1,2,3,4,5,6,7,8,9]
    Blocks = [1,2,3,4,5,6,7,8,9]
    
    # Define boolean variables
    x = {}
    for s in Sections:
        for j in Courses:
            for k in Blocks:
                x[s,j,k] = solver.IntVar(0,1, 'x[%d,%d,%d]' % (s,j,k))

    y = {}
    for i in Students:
        for j in Courses:
            for k in Blocks:
                y[i,j,k] = solver.IntVar(0,1, 'y[%d,%d,%d]' % (i,j,k)) 


    # CONSTRAINT 1: For each course, ensure the correct number of sections are offered.
    for j in Courses:
        for s in Sections:
            if s <= CourseSections[j]:
                solver.Add(sum(x[s,j,k] for k in Blocks) == 1)
            else:
                solver.Add(sum(x[s,j,k] for k in Blocks) == 0)

                
    # CONSTRAINT 2: Two sections of the same course can't be offered in the same block
    for j in Courses:
        for k in Blocks:
            solver.Add(sum(x[s,j,k] for s in Sections) <= 1)


    # CONSTRAINT 3: For each teacher, all of their required courses must occur in separate blocks
    for t in Teachers:
        for k in Blocks:
            solver.Add( sum(x[s,j,k] for s in Sections for j in TeacherCourses[t]) <= 1)              


    # CONSTRAINT 4: Ensure forbidden assignments are not made
    for z in ForbiddenAssignments:
        j = z[0]
        k = z[1]
        for s in Sections:
            solver.Add(x[s,j,k]==0)


    # CONSTRAINT 5: ensure required assignments are made
    for z in RequiredAssignments:
        s = z[0]
        j = z[1]
        k = z[2]
        solver.Add(x[s,j,k]==1)


    # CONSTRAINT 6: No room can be used twice in the same block.
    for p in Courses:
        for q in Courses:
            if p<q and RoomChoices[p]==RoomChoices[q]:
                if len(RoomChoices[p])==1 and RoomChoices[p] != ['General'] and RoomChoices[p] != ['nan']:
                    for k in Blocks:
                        solver.Add( sum(x[s,p,k]+x[s,q,k] for s in Sections) <= 1)


    # CONSTRAINT 7: Due to room constraints, every block can have at most 4 courses from each of
    # the following departments: English, Mathematics, Languages, Science, and Social Studies.
    # These are the exact five departments identified in the DepartmentCourses variable.

    for d in range(5):
        for k in Blocks:
            solver.Add(sum(sum(x[s,j,k] for j in DepartmentCourses[d]) for s in Sections) <= 5)


    # CONSTRAINT 8: Ensure CALC12, APCALA, APCAL12 are all in the same block, with PH12 not
    # being in that block.
    
    j1 = CourseList.index("AP Calculus AB")
    j2 = CourseList.index("AP Calculus BC")
    j3 = CourseList.index("Calculus 12")
    j4 = CourseList.index("Physics 12")
    for k in Blocks:
        solver.Add(x[1,j1,k]==x[1,j2,k])
        solver.Add(x[1,j2,k]==x[1,j3,k])
        solver.Add(x[1,j1,k]+x[1,j4,k]+x[2,j4,k] <= 1)
        solver.Add(x[1,j2,k]+x[1,j4,k]+x[2,j4,k] <= 1)
        solver.Add(x[1,j3,k]+x[1,j4,k]+x[2,j4,k] <= 1)
    
    
            
    # CONSTRAINT 9: Each student takes at most one course per block
    for i in Students:
        for k in Blocks: 
            solver.Add(sum(y[i,j,k] for j in Courses) <= 1)


    # CONSTRAINT 10: No student can take the same course twice       
    for i in Students:
        for j in Courses:
            solver.Add(sum(y[i,j,k] for k in Blocks) <= 1)
               

    # CONSTRAINT 11: No student can take a course in a block when that course isn't offered
    for i in Students:
        for j in Courses:
            for k in Blocks:
                solver.Add(y[i,j,k] <= sum(x[s,j,k] for s in Sections))


    # CONSTRAINT 12: Do not assign course j to a student i if P[i,j]=0
    for i in Students:
        for j in Courses:
            if P[i,j]==0:    
                for k in Blocks:
                    solver.Add(y[i,j,k]==0)


    # CONSTRAINT 13: No course section can exceed its room capacity
    for j in Courses:
        for k in Blocks:
            solver.Add(sum(y[i,j,k] for i in Students) <= RoomLimit[j])   

            
    # CONSTRAINT 14: No student can take StudyBlock and StudyBlock2 on the same day.
    # except students 155 and 234
    j1 = CourseList.index("Study Block")
    j2 = CourseList.index("Study Block2")
    for i in Students:
        if not StudentList[i] in [155,234]:
            solver.Add(sum(y[i,j1,k] + y[i,j2,k] for k in [1,2,3,4]) <= 1)
            solver.Add(sum(y[i,j1,k] + y[i,j2,k] for k in [5,6,7,8,9]) <= 1)
         
            
    # CONSTRAINT 15: At most 30 students can be in a Study Block in any given block
    j1 = CourseList.index("Study Block")
    j2 = CourseList.index("Study Block2")
    for k in Blocks:
        solver.Add(sum(y[i,j1,k] + y[i,j2,k] for i in Students) <= 30)   
    
            
    # CONSTRAINT 16: For all of the x[s,j,k] assignments from XSet, lock in all of them
    # except for some number of course sections (defined by FixedNumber) that we can move 
    # to other blocks to optimize the quality of our timetable.  To do this, we first use the
    # random package to shuffle XSet, and then allow only the first FixedNumber course sections 
    # of our shuffled XSet to be changed.
    
    shuffle(XSet)
    for z in range(FixedNumber, len(XSet)):
        s = XSet[z][0]
        j = XSet[z][1]
        k = XSet[z][2]
        solver.Add(x[s,j,k] == 1)
 

    # CONSTRAINT 17: Add our IEP constraints
    
    for j in IEPcourses:
        for k in Blocks:
            if CourseSections[j] == 2:
                solver.Add(sum(y[i,j,k] for i in Students if IEP[i][j] == 1)
                       <= 0.60 * sum(IEP[_][j] for _ in range(len(IEP)))) 
            if CourseSections[j] == 3:
                solver.Add(sum(y[i,j,k] for i in Students if IEP[i][j] == 1)
                       <= 0.40 * sum(IEP[_][j] for _ in range(len(IEP)))) 
            if CourseSections[j] == 4:
                solver.Add(sum(y[i,j,k] for i in Students if IEP[i][j] == 1)
                       <= 0.31 * sum(IEP[_][j] for _ in range(len(IEP)))) 
            if CourseSections[j] == 5:
                solver.Add(sum(y[i,j,k] for i in Students if IEP[i][j] == 1)
                       <= 0.25 * sum(IEP[_][j] for _ in range(len(IEP)))) 

        
    # CONSTRAINT 18: Add balancing constraints to ensure each course section has roughly the
    # same number of students.  No 2-section course can have more than 54% of the enrolled 
    # students in one section.  Do the same for 3-section, 4-section, and 5-section courses.
    
    
    # NOTE TO ME - change this back to what I had earlier (0.54, 0.36, 0.3, 0.27, 0.25)
    
    for j in Courses:
        for k in Blocks:
            if CourseSections[j]==2:
                solver.Add(sum(y[i,j,k] for i in Students) <= 0.54 * CourseRequestTotal[j])  
            if CourseSections[j]==3:
                if "8." in CourseList[j]:
                    solver.Add(sum(y[i,j,k] for i in Students) <= 0.4 * CourseRequestTotal[j])
                else:
                    solver.Add(sum(y[i,j,k] for i in Students) <= 0.36 * CourseRequestTotal[j])
            if CourseSections[j]==4:
                solver.Add(sum(y[i,j,k] for i in Students) <= 0.265 * CourseRequestTotal[j])
            if CourseSections[j]==5:
                if CourseList[j] == "Guided Study Block":
                    solver.Add(sum(y[i,j,k] for i in Students) <= 0.4 * CourseRequestTotal[j]) 
                elif "8." in CourseList[j]:
                    solver.Add(sum(y[i,j,k] for i in Students) <= 0.24 * CourseRequestTotal[j]) 
                else:
                    solver.Add(sum(y[i,j,k] for i in Students) <= 0.22 * CourseRequestTotal[j]) 
                    
    for k in Blocks:
        j = CourseList.index("Active Living 11/12")
        solver.Add(sum(y[i,j,k] for i in Students) <= 19)
        j = CourseList.index("Pre-Calculus 11")
        solver.Add(sum(y[i,j,k] for i in Students) <= 15)
       
    
    
    # Solve the Integer Linear Program!
    solver.Maximize(solver.Sum(P[i,j]*y[i,j,k]
                   for i in Students for j in Courses for k in Blocks))
    sol = solver.Solve()
    ObjectiveValue = round(solver.Objective().Value())
    
    
    # Generate the new XSet (the master timetable from the perspective of the courses) and the
    # new YSet (the master timetable from the perspective of the students)
    
    XSet=[]
    for s in Sections:
        for j in Courses:
            for k in Blocks:
                if x[s,j,k].solution_value()==1:
                    XSet.append([s,j,k])            
    YSet=[]
    for i in Students:
        for j in Courses:
            for k in Blocks:
                if y[i,j,k].solution_value()==1:
                    YSet.append([i,j,k])
                        
    return [ObjectiveValue, XSet, YSet]

# Pre-load the best timetable found so far

XSet = [[1, 0, 8], [1, 1, 1], [1, 2, 3], [1, 3, 3], [1, 4, 2], [1, 5, 4], [1, 6, 6], [1, 7, 9], [1, 8, 5], [1, 9, 6], [1, 10, 6], [1, 11, 8], [1, 12, 7], [1, 13, 8], [1, 14, 1], [1, 15, 9], [1, 16, 2], [1, 18, 7], [1, 19, 9], [1, 20, 9], [1, 21, 4], [1, 23, 2], [1, 25, 3], [1, 26, 4], [1, 27, 4], [1, 28, 8], [1, 29, 4], [1, 31, 9], [1, 32, 2], [1, 33, 6], [1, 34, 1], [1, 35, 5], [1, 36, 2], [1, 37, 4], [1, 38, 1], [1, 39, 4], [1, 40, 3], [1, 41, 3], [1, 42, 6], [1, 43, 4], [1, 44, 8], [1, 45, 7], [1, 46, 3], [1, 48, 5], [1, 49, 8], [1, 50, 6], [1, 51, 1], [1, 52, 5], [1, 54, 8], [1, 55, 9], [1, 56, 7], [1, 57, 7], [1, 58, 7], [1, 59, 9], [1, 60, 3], [1, 61, 9], [1, 62, 4], [1, 63, 8], [1, 64, 8], [1, 65, 4], [1, 66, 7], [1, 67, 3], [1, 68, 6], [1, 69, 7], [1, 70, 1], [1, 71, 4], [1, 72, 4], [1, 74, 2], [1, 75, 2], [1, 76, 1], [1, 77, 2], [1, 78, 2], [1, 79, 8], [1, 80, 5], [1, 81, 3], [1, 82, 7], [1, 83, 9], [1, 84, 5], [1, 85, 2], [1, 87, 1], [1, 88, 3], [1, 89, 7], [1, 90, 9], [1, 91, 6], [1, 92, 9], [1, 93, 5], [1, 94, 1], [1, 95, 6], [1, 96, 2], [1, 97, 2], [1, 98, 1], [1, 99, 1], [1, 100, 9], [1, 101, 1], [1, 102, 7], [1, 103, 3], [1, 104, 9], [1, 105, 5], [1, 106, 2], [1, 107, 3], [1, 108, 8], [1, 109, 6], [1, 110, 2], [1, 111, 6], [1, 112, 1], [1, 113, 3], [1, 114, 4], [1, 115, 3], [1, 116, 2], [1, 117, 7], [1, 118, 5], [1, 119, 1], [1, 120, 1], [1, 121, 1], [1, 122, 1], [1, 123, 5], [1, 124, 6], [1, 125, 7], [1, 126, 4], [1, 127, 3], [1, 128, 2], [2, 0, 2], [2, 8, 7], [2, 9, 9], [2, 14, 6], [2, 16, 4], [2, 19, 4], [2, 21, 7], [2, 23, 5], [2, 26, 2], [2, 28, 7], [2, 29, 6], [2, 31, 2], [2, 38, 4], [2, 45, 9], [2, 46, 5], [2, 48, 3], [2, 49, 5], [2, 50, 8], [2, 51, 5], [2, 54, 4], [2, 55, 3], [2, 58, 9], [2, 59, 6], [2, 60, 8], [2, 61, 3], [2, 62, 2], [2, 67, 8], [2, 68, 2], [2, 69, 9], [2, 72, 7], [2, 75, 1], [2, 76, 9], [2, 77, 6], [2, 78, 3], [2, 79, 7], [2, 80, 3], [2, 88, 4], [2, 89, 8], [2, 90, 8], [2, 91, 9], [2, 100, 7], [2, 101, 9], [2, 102, 1], [2, 103, 5], [2, 105, 3], [2, 106, 9], [2, 107, 7], [2, 109, 4], [2, 110, 7], [2, 111, 4], [2, 112, 5], [2, 113, 4], [2, 119, 2], [2, 120, 2], [2, 121, 2], [2, 124, 1], [2, 126, 1], [3, 19, 8], [3, 26, 5], [3, 28, 6], [3, 48, 7], [3, 49, 3], [3, 50, 9], [3, 54, 2], [3, 59, 4], [3, 67, 9], [3, 69, 4], [3, 72, 8], [3, 75, 3], [3, 77, 3], [3, 79, 1], [3, 89, 4], [3, 102, 6], [3, 103, 8], [3, 105, 4], [3, 107, 2], [3, 109, 8], [3, 110, 5], [3, 111, 8], [3, 112, 6], [3, 113, 8], [3, 119, 3], [3, 120, 3], [3, 121, 3], [4, 19, 7], [4, 26, 9], [4, 48, 6], [4, 49, 7], [4, 72, 2], [4, 102, 2], [4, 103, 1], [4, 107, 9], [4, 109, 1], [4, 110, 8], [4, 111, 5], [4, 112, 9], [4, 113, 2], [4, 119, 4], [4, 120, 4], [4, 121, 4], [5, 19, 3], [5, 72, 3], [5, 102, 5], [5, 109, 3], [5, 111, 9], [5, 112, 4], [5, 113, 6], [5, 119, 5], [5, 120, 5], [5, 121, 5], [6, 119, 6], [6, 120, 6], [6, 121, 6], [7, 119, 7], [7, 120, 7], [7, 121, 7], [8, 119, 8], [8, 120, 8], [8, 121, 8], [9, 119, 9], [9, 120, 9], [9, 121, 9]]


# Moved Pre-Calc 12 from 1C to 2B
XSet.remove([1,103,3])
XSet.append([1,103,6])


# Moved Socials 8 from 2E to 2D

XSet.remove([4,112,9])
XSet.append([4,112,8])

# Now use the Initial Timetable (XSet) of just the course/section assignments to blocks
# to generate the YSet, the optimal assignment of students to courses and blocks for this timetable.

start_time = time.time()
FirstIteration = HillClimber(XSet, 0)
ObjectiveValue = FirstIteration[0]
XSet = FirstIteration[1]
YSet = FirstIteration[2]
solving_time = round(time.time() - start_time)

print("Iteration 0 complete in", solving_time, "seconds with", ObjectiveValue, "points and", 
      len(YSet), "student requests satisfied")

file = open('CS5100_XSet.txt', 'w')
file.write(str(XSet))
file.close()
file = open('CS5100_YSet.txt', 'w')
file.write(str(YSet))
file.close()

# Generate the statistics for our Master Timetable, to see how well our timetable
# assigned students to their requested courses.

Courses = range(m)
TotalRequests = [0 for g in range(13)]
TotalAssignments = [0 for g in range(13)]

for g in range(13):
    for i in StudentsPerGrade[g]:
        for j in Courses:
            if P[i][j]>0: 
                TotalRequests[g] += 1
    for y in YSet:
        i = y[0]
        j = y[1]
        if i in StudentsPerGrade[g]:
            TotalAssignments[g] += 1

print("Statistics for Our Master Timetable for the entire WPGA Senior School:")
print(sum(TotalAssignments), "out of", sum(TotalRequests), "total courses satisfied:",
          round(100*sum(TotalAssignments)/sum(TotalRequests),2), "percent")

for g in [12,11,10,9,8]:
    print("")
    print("Results for Grade", g, "Students")
    print(TotalAssignments[g], "out of", TotalRequests[g], "total courses satisfied:",
          round(100*TotalAssignments[g]/TotalRequests[g],2), "percent")

MissedCourses = []
for i in range(n):
    for j in StudentChoices[i]:
        if P[i][j]>0:
            MissedCourses.append([j,i])
        
for i in range(len(YSet)):
    if P[YSet[i][0]][YSet[i][1]]>0:
        if [YSet[i][1], YSet[i][0]] in MissedCourses:
            MissedCourses.remove([YSet[i][1], YSet[i][0]])
    
MissedCourses.sort()

print("Total Missed Courses:", len(MissedCourses))
print("")
BlockList = ["","1A","1B","1C","1D","2A","2B","2C","2D","2E"]
for mypair in MissedCourses:
    
    j = mypair[0]
    offered = ""
    for x in XSet:
        if x[1]==j:
            if offered == "": offered += BlockList[x[2]]
            else: offered += "/" + BlockList[x[2]]
    
    #if mypair[1] in StudentsPerGrade[12]: 
    #    print("Grade 12 student", StudentList[mypair[1]], "missed", CourseList[mypair[0]], "offered in block", offered)
    
    #if mypair[1] in StudentsPerGrade[11]: 
    #    print("Grade 11 student", StudentList[mypair[1]], "missed", CourseList[mypair[0]], "offered in block", offered)
        
    #if mypair[1] in StudentsPerGrade[10]: 
    #    print("Grade 10 student", StudentList[mypair[1]], "missed", CourseList[mypair[0]], "offered in block", offered)
        
    #if mypair[1] in StudentsPerGrade[9]: 
    #    print("Grade 9 student", StudentList[mypair[1]], "missed", CourseList[mypair[0]], "offered in block", offered)
        
    if mypair[1] in StudentsPerGrade[8]: 
        print("Grade 8 student", StudentList[mypair[1]], "missed", CourseList[mypair[0]], "offered in block", offered)

# Use this code to double-check the enrollment numbers to ensure the multi-section
# courses have roughly the same number of students.
for j in range(m):
    ycount = [0,0,0,0,0,0,0,0,0]
    for k in [1,2,3,4,5,6,7,8,9]:
        for y in YSet:
            if y[1]==j and y[2]==k: ycount[k-1] += 1
    Missing = CourseRequestTotal[j] - sum(ycount)
    if CourseSections[j]==5:
        print(ycount, CourseList[j], "has", CourseRequestTotal[j], "enrolled students and",
              Missing, "unenrolled students")

for j in IEPcourses:
    ycount = [0,0,0,0,0,0,0,0,0]
    for k in [1,2,3,4,5,6,7,8,9]:
        for y in YSet:
            if y[1] == j and y[2] == k:
                if IEP[y[0]][j] == 1: ycount[k-1] += 1
    if CourseSections[j]==4:
        print(ycount, CourseList[j], "has", sum(ycount), "IEP students split into", CourseSections[j], "sections")

for j in range(m):
    ycount = [[0,0],[0,0],[0,0],[0,0],[0,0],[0,0],[0,0],[0,0],[0,0]]
    for k in [1,2,3,4,5,6,7,8,9]:
        for y in YSet:
            if y[1] == j and y[2] == k:
                if GenderInfo[y[0]] == 1: ycount[k-1][1] += 1
                if GenderInfo[y[0]] == 0: ycount[k-1][0] += 1
    while [0,0] in ycount:
        ycount.remove([0,0])
    if CourseSections[j]>1 and CourseSections[j]<6:
        print(ycount, CourseList[j], "has this Female/Male split")

# Output the Master Timetable as two Excel sheets: one from the perspective of the courses
# and one from the perspective of the students

OurColumns = ["Course Name", "Section", "Block", "Enrollment", "Capacity", "Room", "Teachers"]
Blocks = ["0", "1A", "1B", "1C", "1D", "2A", "2B", "2C", "2D", "2E"]
M = []
for x in range(len(XSet)):
    s = XSet[x][0]
    j = XSet[x][1]
    k = XSet[x][2]
    count=0
    for y in range(len(YSet)):
        if YSet[y][1]==j and YSet[y][2]==k:
            count+=1
    
    M += [[CourseList[j], 'Section '+str(s), Blocks[k], count, 
           RoomLimit[j], RoomChoices[j], PossibleTeachers[j]]]
    
FinalMatrix = pd.DataFrame(M, columns=OurColumns)
FinalMatrix.to_csv("WPGA Optimal Timetable (Courses).csv", index = False)



OurColumns = ["StudentID", "Student Grade", "Course Title", "Course Code", 
              "Preference", "Block"]

Blocks = ["0", "1A", "1B", "1C", "1D", "2A", "2B", "2C", "2D", "2E"]
M = []
for x in range(len(InputInfo)):
    Response = ""
    i = StudentList.index(InputInfo[x][31])
    j = CourseList.index(InputInfo[x][35]) 
    if CourseSections[j] == 0:
        Response = "Not Scheduled"
    else:       
        for y in YSet:
            if y[0]==i and y[1]==j: 
                Response = Blocks[y[2]]
                
    if Response == "":
        Response = "FAIL"
        
    M += [[InputInfo[x][31], InputInfo[x][34], InputInfo[x][35], InputInfo[x][36],  
          InputInfo[x][37], Response]]
    
FinalMatrix = pd.DataFrame(M, columns=OurColumns)
FinalMatrix.to_csv("WPGA Optimal Timetable (Students).csv", index = False)

for k in [1,2,3,4,5,6,7,8,9]:
    xcount=0
    ycount=[0,0,0,0,0]
    for x in XSet:
        if x[2]==k:
            xcount+=1
    for y in YSet:
        if y[2]==k:
            if y[0] in StudentsPerGrade[8]: ycount[0]+=1
            if y[0] in StudentsPerGrade[9]: ycount[1]+=1
            if y[0] in StudentsPerGrade[10]: ycount[2]+=1
            if y[0] in StudentsPerGrade[11]: ycount[3]+=1
            if y[0] in StudentsPerGrade[12]: ycount[4]+=1
    print("Block", BlockList[k], "has", ycount, "students enrolled in", xcount, "courses")

for d in range(5):
    for k in range(10):
        count = 0
        for x in XSet:
            if x[1] in DepartmentCourses[d]:
                if x[2]==k: 
                    count+=1
        if count>=5:
            print("WARNING: department", Departments[d], "has", count, "courses offered in block", k)