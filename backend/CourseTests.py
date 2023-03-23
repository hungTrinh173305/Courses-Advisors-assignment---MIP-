import unittest
from datetime import time
from Course import Course

"""
Author: Riley Spagnuolo
This is a series of tests for Course.py 
"""
class CourseTests(unittest.TestCase):

    def testCourseCode(self):
        course_1 = Course("AAH-103-01 Intro Eur Painting & Sculpture", 0, 0, ["Monday", "Wednesday", "Friday"])
        self.assertEquals("AAH-103-01",course_1.getCourseCode())

        course_2 = Course("*BIO-103-02 Diversity of Life W/Lab", 0, 0, ["Monday", "Wednesday", "Friday"])
        self.assertEquals("BIO-103-02",course_2.getCourseCode())

        course_3 = Course("BIO-103L-01 Diversity of Life Lab", 0, 0, ["Monday", "Wednesday", "Friday"])
        self.assertEquals("BIO-103L-01",course_3.getCourseCode())

        course_4 = Course("BIO-103L-02 Diversity of Life Lab", 0, 0, ["Monday", "Wednesday", "Friday"])
        print(course_4.getCourseCode())
        
    def testOverlappingTimes(self):
        time1 = time(hour = 14, minute = 30)
        time2 = time(hour = 13, minute = 45)
        time3 = time(hour = 15, minute = 45)
        time4 = time(hour = 11, minute = 45)
        self.assertTrue(Course.isConflict(self,time1 = time1,startTime=time2,endTime=time3)) #Tests if there is a conflict
        self.assertFalse(Course.isConflict(self,time1 = time4,startTime=time2,endTime=time3)) #no conflict

    def testOverlappingCourses(self):
        time1 = time(hour=11, minute=30)
        time2 = time(hour=12, minute=45)
        time3 = time(hour=13, minute=45)
        time4 = time(hour=14, minute=45)
        name = "csc-100"
        department = "computer science"
        days1 = ["monday","wednesday","friday"]
        days2 = ["tuesday","thursday"]
        course1 = Course(name,department,time1,time3,days = days1)
        course2 = Course(name,department,time2,time4,days = days1)
        course3 = Course(name,department,time1,time2,days = days1)
        course4 = Course(name,department,time3,time4,days = days1)
        course5 = Course(name,department,time1,time2,days = days2)
        self.assertTrue(hasOverlap(course1,course2))
        self.assertFalse(hasOverlap(course3,course4))
        self.assertFalse(hasOverlap(course4,course5))

if __name__ == "__main__":
    unittest.main()