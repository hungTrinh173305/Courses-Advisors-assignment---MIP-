
class Student:

    def __init__(self, name, id, assignedAdvisor=None, max_courses=3):
        self.name = name
        self.id = id

        self.student_detail_dict = {}

        self.assignedAdvisor = assignedAdvisor

        self.courses = []
        self.max_courses = max_courses

    def getName(self):
        """
        returns student's name.
        """
        return self.name

    def getID(self):
        """
        returns student id.
        """
        return self.id

    def set_student_detail(self, key, value):
        """
        @param key: the attribute to set.
        @param value: the value to set the attribute to.
        sets a specific student detail to a given value.
        """
        self.student_detail_dict[key] = value

    def get_student_detail(self, key):
        """
        @param key: the attribute to return.
        returns a specific student detail.
        """
        return self.student_detail_dict[key]

    def getAdvisor(self):
        """
        returns the advisor assigned to this student.
        """
        return self.assignedAdvisor

    def hasAdvisor(self):
        """
        returns a boolean representing whether this student has been assigned an advisor.
        """
        return self.assignedAdvisor is not None

    def assignAdvisor(self, advisor):
        """
        assigns a given advisor to this student. if the student previously had an assigned advisor, they are removed
        from that advisor's list of advisees.
        """
        if self.assignedAdvisor is not None:
            self.assignedAdvisor.removeAdvisee(self)

        self.assignedAdvisor = advisor
        advisor.assignAdvisee(self)

    def addCourse(self, course):
        """
        Adds a course to the list of courses that a student is taking.
        If a student is already taking this course then does nothing.
        """
        if course not in self.courses:
            self.courses.append(course)

    def removeCourse(self, course):
        """
        Removes a course from the list of courses that a student is taking.
        If a student is not taking this course then does nothing.
        """
        if course in self.courses:
            self.courses.remove(course)

    def countCourses(self):
        return len(self.courses)

    def getMaxCourses(self):
        """
        returns the maximum number of students allowed to take this course
        """
        return self.max_courses

    def getOpenSlots(self):
        """
        returns the number of "slots" available for this course, i.e. the maximum number of students minus the number
        of students already taking this course
        """
        return self.getMaxCourses() - len(self.courses)

    def slotsRemaining(self):
        """
        returns a boolean representing whether this course is capable of taking another student
        """
        return self.getOpenSlots() > 0

    def getCourses(self):
        """
        returns a list of courses that this student is enrolled in
        """
        return self.courses

    def _check_courses_preferred(self, key):
        course = self.get_student_detail(key)
        if course[0] == '*':
            major_interest_list = []
            major_1 = self.get_student_detail("Maj_Int_1")
            major_2 = self.get_student_detail("Maj_Int_2")
            major_3 = self.get_student_detail("Maj_Int_3")
            major_4 = self.get_student_detail("Maj_Int_4")
            major_interest_list.append(major_1)
            major_interest_list.append(major_2)
            major_interest_list.append(major_3)
            major_interest_list.append(major_4)
            return major_interest_list
        else:
            if ("-" not in course):
                return course[:3]
            else: 
                return course

    def _format_ECO101_in_preferred_courses(self, pref_list):
        return_list = []
        for aclass in pref_list:
            if aclass == 'ECO- 101 Introduction to Economics':
                aclass = 'ECO-101 Introduction to Economics'
                return_list.append(aclass)
            else:
                return_list.append(aclass)
        return return_list

    def get_preferred_courses(self):
        """
        getting a student preferred courses and FYPs matching the data representation needed for the scoring function.
        """
        pref_list = []
        fyp_list = []
        pref_1 = self._check_courses_preferred("Preferred Course_1")
        pref_3 = self._check_courses_preferred("Preferred Course_3")
        pref_2 = self._check_courses_preferred("Preferred Course_2")
        pref_4 = self._check_courses_preferred("Preferred Course_4")
        pref_5 = self._check_courses_preferred("Preferred Course_5")
        pref_6 = self._check_courses_preferred("Preferred Course_6")
        pref_7 = self._check_courses_preferred("Preferred Course_7")

        fyp_1 = self.get_student_detail("FYP_Preferred_FALL 1")
        fyp_2 = self.get_student_detail("FYP_Preferred_FALL 2")
        fyp_3 = self.get_student_detail("FYP_Preferred_FALL 3")
        fyp_4 = self.get_student_detail("FYP_Preferred_FALL 4")
        fyp_5 = self.get_student_detail("FYP_Preferred_FALL 5")
        fyp_6 = self.get_student_detail("FYP_Preferred_Reason & Passion")

        pref_list.append(pref_1)
        pref_list.append(pref_2)
        pref_list.append(pref_3)
        pref_list.append(pref_4)
        pref_list.append(pref_5)
        pref_list.append(pref_6)
        pref_list.append(pref_7)

        fyp_list.append(fyp_1)
        fyp_list.append(fyp_2)
        fyp_list.append(fyp_3)
        fyp_list.append(fyp_4)
        fyp_list.append(fyp_5)
        fyp_list.append(fyp_6)

        pref_list.append(fyp_list)
        return_l = self._format_ECO101_in_preferred_courses((pref_list))
        return return_l
