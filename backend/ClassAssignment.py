import sys
sys.path.append("../")

import backend.ExcelManagement as em
from backend.Student import Student
from backend.Advisor import Advisor  # need?
from backend.Course import Course


class classAssignment:

    def __init__(self):

        self.courses = []
        self.students = []

    def add_student(self, student):
        """
        adds a student to the class assignment manager.
        """
        self.students.append(student)

    def add_course(self, course: Course):
        """
        adds a course to the class assignment manager
        """
        self.courses.append(course)

    def assign(self, student, course):
        """
        assigns a student to a course s/t both the student and the course are aware of the assignment
        @param student: the student to assign.
        @param course: the course the student is being assigned to.
        """
        student.addCourse(course)
        course.addStudent(student)

    def get_student(self, student_ID):
        """
        returns a Student object if one with a given student ID exists within the advisor assignment manager.
        @param student_ID: id number of the student you are looking for.
        """
        for student in self.students:
            if student.getID() == student_ID:
                return student
        return None

    def get_course_by_name(self, course_name):
        """
        returns a course based on the course name
        """
        for course in self.courses:
            if course.getName() == course_name:
                return course
        return None

    def get_course_by_code(self, department, number, section):
        for course in self.courses:
            if course.getDepartment() == department:
                if course.getNumber() == number:
                    if course.getSection() == section:
                        return course
        return None

    def get_students(self):
        """
        returns a list of all the students that exist within the course assignment manager.
        """
        return self.students

    def get_courses(self):
        """
        returns a list of all the courses that exist within the course assignment manager
        """
        return self.courses

    def output(self):
        """
        prints out the state of the manager. prints each student's name followed by the name of the advisor that student
        is assigned to
        """
        for student in self.students:
            print(student.getName(), "->", student.getCourses())


def create_manager():
    """
    creates an instance of the classAssignment class based on the information in the relevant workbook.
    reads in all the courses available and all the students from the spreadsheet.
    """
    manager = classAssignment()

    class_info = em.read_workbook("COURSES OFFERED")  # make this global / constant?
    found_last = False
    i = 0
    while not found_last:
        if class_info[i][0] is None:
            class_info = class_info[:i + 1]  # does this include empty row? start "course?" or does em handle this?
            # this could be an error if a value accidentally gets deleted & script sees it as last
            found_last = True
        i += 1

    for row in class_info:
        if row[0] is None:  # if there's no course title, ignore
            pass  # we might be able to remove these rows from data & circumvent passes
        # i.e. remove_rows('-'), remove_rows('--'), etc. need to make sure we are passing a COPY first.
        # this case never happens because of the while loop
        elif row[0][:9] == "CANCELLED":  # ignore any classes marked cancelled 
            # requires consistent notation of cancelled classes
            pass
        elif "-" not in row[0]:  # ignore any rows that don't have dashes in them (these rows are not valid courses)
            pass
        elif "--" in row[0]:
            pass
        else:
            schedule = {}
            if row[13] == "M":
                schedule["Monday"] = (str(row[18]), str(row[19]))
            if row[14] == "W":
                schedule["Wednesday"] = (str(row[20]), str(row[21]))
            if row[15] == "F":
                schedule["Friday"] = (str(row[22]), str(row[23]))
            if row[16] == "T":
                schedule["Tuesday"] = (str(row[24]), str(row[25]))
            if row[17] == "TH":
                schedule["Thursday"] = (str(row[26]), str(row[27]))
            manager.add_course(Course(row[0], schedule, int(row[7]), int(row[8])))

    student_detail_keys = em.read_workbook_row("Main Tab", 0)
    student_info = em.read_workbook("Main Tab")[1:]
    found_last = False
    i = 0
    while not found_last:
        if student_info[i][0] is None:
            student_info = student_info[:i + 1]
            found_last = True
        i += 1

    for student_as_list in student_info:
        student_name = student_as_list[4]
        student_id = student_as_list[0]
        student_as_object = Student(student_name, student_id)

        for i in range(len(student_detail_keys)):
            key = student_detail_keys[i]
            val = student_as_list[i]
            student_as_object.set_student_detail(key, val)

        manager.add_student(student_as_object)

    return manager


if __name__ == '__main__':
    class_assignment_manager = create_manager()
    for course in class_assignment_manager.get_courses():
        print(course.getCourseCode())
