class Advisor:

    def __init__(self, name, maxStudents=8, currStudents=0):
        self.name = name
        self.advisees = []
        self.maxStudents = maxStudents
        self.currStudents = currStudents
        self.department = ""

    def getName(self):
        """
        returns advisor name.
        """
        return self.name

    def getCurrStudents(self):
        return self.currStudents
    
    def getMaxStudents(self):
        """
        returns the maximum number of students allowed to be assigned to this advisor.
        """
        return self.maxStudents

    def getOpenSlots(self):
        """
        returns the number of "slots" available for this advisor, i.e. the maximum number of students minus the number
        of students already assigned to this advisor.
        """
        return self.getMaxStudents() - len(self.advisees)

    def slotsRemaining(self):
        """
        returns a boolean representing whether this advisor is capable of being assigned another student.
        """
        return self.getOpenSlots() > 0

    def assignAdvisee(self, student):
        """
        assigns a specific student to this advisor. if the student is already assigned to this advisor, does nothing.
        @param student: the student to assign to this advisor.
        """
        if student not in self.advisees:
            self.advisees.append(student)

    def removeAdvisee(self, student):
        """
        removes a specific student from this advisor's list of advisees. if the student was not previously in this
        advisor's list of advisees, does nothing.
        @param student: the student to remove.
        """
        if student in self.advisees:
            self.advisees.remove(student)
            
    def setDepartment(self, department):
        """
        sets this advisor's department.
        @param department: the department to set.
        """
        self.department = department
        
    def getDepartment(self):
        """
        returns this advisor's department.
        """
        return self.department
