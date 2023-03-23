import sys
sys.path.append("../")
import backend.ExcelManagement as em
from backend.Student import Student
from backend.Advisor import Advisor

class advisorAssignment:

    def __init__(self):

        self.advisors = []
        self.students = []
        self.mapping = self.make_mapping()

    def make_mapping(self, worksheet="Sheet1", workbook="../input_files/major_to_dept_matching.xlsx"):
        """
        creates a mapping between student majors and advisor departments based on an Excel workbook with the relevant
        information.
        @param worksheet: the name of the worksheet this information is stored in. defaults to Sheet1, the default
        worksheet name in Excel.
        @param workbook: the name of the workbook (i.e. Excel file) this information is stored in. Defaults to
        major_to_dept_matching.xlsx, the file in which the data is stored.
        """
        mapping = []
        major_to_dept_spreadsheet = em.read_workbook(worksheet, workbook)
        for i in range(1, len(major_to_dept_spreadsheet)):  # skip first row - header row
            to_append = []
            for cell in major_to_dept_spreadsheet[i]:
                if cell is not None:
                    to_append.append(cell)
            mapping.append(to_append)
        return mapping

    def get_mapping(self):
        """
        returns the mapping as a list of lists.
        """
        return self.mapping

    def add_advisor(self, advisor):
        """
        adds an advisor to the advisor assignment manager.
        """
        self.advisors.append(advisor)

    def add_student(self, student):
        """
        adds a student to the advisor assignment manager.
        """
        self.students.append(student)

    def assign(self, student, advisor):
        """
        assigns a student to an advisor.
        @param student: the student to assign.
        @param advisor: the advisor for the student to be assigned to.
        """
        student.assignAdvisor(advisor)

    def get_advisor(self, advisor_name):
        """
        returns an Advisor object if one with a given name exists within the advisor assignment manager.
        @param advisor_name: the name of the advisor to return.
        """
        for advisor in self.advisors:
            if advisor.getName() == advisor_name:
                return advisor
        return None

    def get_student(self, student_ID):
        """
        returns a Student object if one with a given student ID exists within the advisor assignment manager.
        @param student_ID: the name of the advisor to return.
        """
        for student in self.students:
            if student.getID() == student_ID:
                return student
        return None

    def get_students(self):
        """
        returns a list of all the students that exist within the advisor assignment manager.
        """
        return self.students

    def get_advisors(self):
        """
        returns a list of all the advisors that exist within the advisor assignment manager.
        """
        return self.advisors

    def output(self):
        """
        prints out the state of the manager. prints each student's name followed by the name of the advisor that student
        is assigned to
        """
        for student in self.students:
            print(student.getName(), "->", student.getAdvisor().getName())

    def maps_to(self, student_major, advisor_department):
        for i in range(len(self.mapping)):
            row = self.mapping[i]
            if student_major == row[0]:
                if advisor_department in row:
                    return True
        return False


    def score_assignment(self, student: Student, advisor: Advisor):
        if advisor is None or student is None:
            return 0
        advisor_department = advisor.getDepartment()
        student_is_declared = student.get_student_detail("Maj_Declare in 1st Interest?")
        if student_is_declared == "Yes":
            major = student.get_student_detail("MAJOR")
            if self.maps_to(major, advisor_department):
                return 10

        interest_1 = student.get_student_detail("Maj_Int_1")
        interest_2 = student.get_student_detail("Maj_Int_2")
        interest_3 = student.get_student_detail("Maj_Int_3")
        interest_4 = student.get_student_detail("Maj_Int_4")

        if self.maps_to(interest_1, advisor_department):
            return 5
        elif self.maps_to(interest_2, advisor_department):
            return 4
        elif self.maps_to(interest_3, advisor_department):
            return 3
        elif self.maps_to(interest_4, advisor_department):
            return 2
        else:
            return 1


def create_manager():
    """
    creates an instance of the advisorAssignment class based on the information in the relevant workbook.
    """
    manager = advisorAssignment()
    advisor_info = em.read_workbook("Advisor Loading")[1:]

    for advisor_as_list in advisor_info:
        advisor_name = advisor_as_list[0]
        curr = advisor_as_list[8]
        if curr is None:
            curr = 0
        limit = advisor_as_list[12]
        advisor_as_object = Advisor(advisor_name,
                                    maxStudents=int(limit),
                                    currStudents=int(curr))

        department = advisor_as_list[4]
        advisor_as_object.setDepartment(department)

        
        manager.add_advisor(advisor_as_object)

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
    adv_asgmt = advisorAssignment()
    # for line in adv_asgmt.get_mapping():
    #     print(line)
    print(adv_asgmt.maps_to("Physics", "Physics and Astronomy"))
    # manager = create_manager()  # this takes a while to run and gives a data validation extension error(?)
    # print(manager.get_mapping())
