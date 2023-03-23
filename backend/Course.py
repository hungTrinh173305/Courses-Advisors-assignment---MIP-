from datetime import time
class Course:

    def __init__(self, title, schedule, currStudents, maxStudents, department=None, number=None, section=None):
        """
        @param title: the course title - i.e. "AAH-103-01 Intro Eur Painting & Sculpture"
        @param department: the three letter course prefix - i.e. "AAH" - if not specified, will be inferred by course title
        @param number: the three number course code - i.e. "103" - if not specified, will be inferred by course title
        @param section: the two number section code - i.e. "01" - if not specified, will be inferred by course title
        @param schedule: A dictionary of days mapping day of week to pair of a start and end time.
        """
        self.title = title  # full name of the course

        split_name = self.title.split("-")  # split course name by dashes
        self.isALabCourse = (len(split_name[1]) == 4) and split_name[1][-1] == "L"
        # if this course is a lab, there will be four characters between the dashes.
        self.hasALabCourse = len(split_name[0]) == 4
        # if this course has a lab, there will be four characters before the first dash.

        if department is None:
            self.department = split_name[0][-3:]  # infer department from course title
        else:
            self.department = department

        if number is None:
            self.number = split_name[1]  # infer number from course title
        else:
            self.number = number

        if section is None:
            self.section = split_name[2][:2]  # infer section from course title
        else:
            self.section = section

        self.schedule = schedule

        self.students = []  # list of students in this course
        self.currStudents = currStudents # non-first year students in course
        self.maxStudents = maxStudents  # maximum number of students allowed to enroll in this course

    def getName(self):
        """
        returns course name
        """
        return self.title

    def getDepartment(self):
        """
        returns the three letter course prefix
        """
        return self.department

    def getNumber(self):
        """
        returns the three number course number code
        """
        return self.number

    def getSection(self):
        """
        returns the two number section code
        """
        return self.section

    def countStudents(self):
        return len(self.students) + self.currStudents

    def getMaxStudents(self):
        """
        returns the maximum number of students allowed to take this course
        """
        return self.maxStudents

    def getOpenSlots(self):
        """
        returns the number of remaining seats available for this course, i.e. the maximum number of students minus the number
        of students already taking this course
        """
        return self.getMaxStudents() - self.countStudents()

    def slotsRemaining(self):
        """
        returns a boolean representing whether this course is capable of taking another student
        """
        return self.getOpenSlots() > 0

    def addStudent(self, student):
        """
        assigns a specific student to this course. if the student is already taking the course, does nothing.
        @param student: the student to take this course
        """
        if student not in self.students:
            self.students.append(student)

    def removeStudent(self, student):
        """
        removes a specific student from the list of students taking the course. if the student is not taking this course then does nothing
        @param student: the student to remove.
        """
        if student in self.students:
            self.students.remove(student)

    def getDays(self):
        """
        returns a list of days on which the course meets
        """
        return self.schedule.keys()

    def hasDay(self, day):
        """
        returns a boolean value representing whether the course meets on a given day
        @param day: the day to check whether this course meets on
        """
        return day in self.schedule.keys()

    def addDay(self,day,startTime,endTime):
        """
        Adds a day to this course
        @param day: the day to add
        @param startTime: the start time
        @param endTime: the ending time
        """
        self.schedule[day] = [startTime,endTime]


    def removeDay(self,day):
        """
        Removes a day from this course
        @param day: the day to remove
        """
        schedule.pop(day)

    def getStartTime(self,day):
        """
        returns the start time on a day
        @param day - the day to get the start time
        """
        return self.schedule.get(day)[0]

    def getEndTime(self,day):
        """
        returns the end time of this course
        @param day - the day to get the start time
        """
        return self.schedule.get(day)[1]

    def hasOverlap(self, other):
        """
        returns a boolean value representing whether this course overlaps with another given course.
        @param other: another course, which will be checked for time conflicts.
        """
        for day in self.getDays():
            if day in other.getDays():
                if self.isConflict(self.schedule[day], other.schedule[day]):
                    return True
        return False

    def isConflict(self, time1, time2):
        """
        Helper function for hasOverlap
        @param time1: tuple of interval to check for an overlap
        @param time2: tuple of interval to check for an overlap
        """
        (s1, e1) = time1
        (s2, e2) = time2
        
        return ((s1 <= s2 <= e1) or (s1 <= e2 <= e1)
                or (s2 <= s1 <= e2) or (s2 <= e1 <= e2))

    def isLab(self):
        """
        returns boolean whether this course is a lab
        """
        return self.isALabCourse

    def hasLab(self):
        """
        returns boolean whether this course has a lab
        """
        return self.hasALabCourse

    def isLabOf(self, other):
        """
        returns whether this course is a lab of another given course
        @param other: the other given course
        """
        if not self.isLab():
            return False
        elif not other.hasLab():
            return False
        else:
            return self.getNumber()[:3] == other.getNumber() and self.getDepartment() == other.getDepartment()

    def getCourseCode(self):
        """
        returns the course code as a string
        """
        return self.getDepartment() + "-" + self.getNumber() + "-" + self.getSection()

    def isDifferentSectionOfSameCourse(self, other):
        """
        returns a boolean representing whether this course and a given other course are two different sections of the
        same course
        @param other: the other course to test against
        """
        return self.getDepartment() == other.getDepartment() and self.getNumber() == other.getNumber() and \
            self.getSection() != other.getSection()

    def __str__(self):
        return self.getName()
