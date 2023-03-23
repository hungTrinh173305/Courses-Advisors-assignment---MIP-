import sys
sys.path.append("../")

from pprint import pprint
import backend.ClassAssignment as ClassAssignment
import backend.ExcelManagement as em
import backend.Student as Student
import backend.Advisor as Advisor
from backend.MappingMajorsToCourses import map_courses_to_major


def is_fit_maj(major_list, classes, major_dict, weight=[3, 1, 0.5, 0.1]):
    # this is an example, need data to fill in
    i = 0
    for major in major_list:
        if major is not None and major != 'Interdepartmental (combining the areas of study listed for the next two questions)':
            classes_belong = major_dict[major]
            for name in classes_belong:
                if classes in name:
                    return weight[i]
        i += 1
    # use if statement like this to check each major
    return False


def is_fit_FYI(FYP_list, classes, score=[36, 25, 16, 9, 4, 1]):
    # currerently working on with.
    # if the list is empty, return false
    if not FYP_list:
        return False
    i = 0

    for fyp in FYP_list:
        if fyp is not None:
            # handle some class like "100-15/16 Music & Politics"
            if (fyp[6] == "/" and
                (fyp[:6] == classes[4:10]
                 or (fyp[:4] + fyp[7:9]) == classes[4:10])):
                    return score[i]
            # Check that the course titles match, ignoring stuff at
            # beginning.
            if " ".join((fyp.split("-")[1]).split()[1:]) == classes[11:]:
                return score[i]
            if fyp[:6] == classes[4:10]:
                return score[i]
            
        i += 1
            
    return False


def is_fit_course(prefered, classes):
    # check if currerent class is prefered class
    if prefered[:7] == classes[:7]:
        return True
    else:
        return False


def is_fit_field(field, classes):
    # check if the current class is preferred field
    if field[:3] == classes[:3]:
        return True
    else:
        return False


def score_assignment(prefer_class_data, assigned_class_data, major_dict):
    # take the prefered class list and assigned class list
    # return a data list including ID, match score and number of class
    data = []
    if len(prefer_class_data) != len(assigned_class_data):
        print(len(prefer_class_data), len(assigned_class_data))
    
    for i in range(len(assigned_class_data)):
        classlist = assigned_class_data[i]
        preferedlist = prefer_class_data[i]
        student_ID = classlist[0]  # get student ID
        classlist = classlist[1:]
        score = 0
        # print(preferedlist[:-1])
        # print(preferedlist[-1:])
        # FYPlist = preferedlist.pop() # this will extract FYP prefer list
        for classes in classlist:
            score = score + score_assignment_2(preferedlist, classes, major_dict)
        data.append([student_ID, score])
    return data


def score_assignment_2(prefer_classes, input_class, major_dict, score_list=[49, 36, 25, 16, 9, 4, 1]):
    # Prototype
    # Improved scoring function, take a class and a list of preferred class as input, return a score.
    # ==== Here is an example of prefer_classes:
    # ['ESC-100 Exploring Engineering'(course), 'MTH'(general field), 'GER',
    # ['ER (Undecided Engineering)', 'Undecided Engineering', 'Mechanical Engineering', 'Physics (Lab Science)'] (list of major), 'CHM',
    # ['ER (Undecided Engineering)', 'Undecided Engineering', 'Mechanical Engineering', 'Physics (Lab Science)'],
    # ['ER (Undecided Engineering)', 'Undecided Engineering', 'Mechanical Engineering', 'Physics (Lab Science)'],
    # ['100-02 Humans & Nonhumans', '100-15/16 Music & Politics', '100-10 Resources & Env Justice'](followed by a FYP list)]
    if input_class[0] == '*':
        input_class = input_class[1:]
    if len(input_class.split("-")) >= 2 and ("L" in ((input_class.split("-"))[1])):
        # Make all lab score 0.
        return 0

    preferedlist = prefer_classes[:-1]
    score = 0
    i = 0
    FYPlist = prefer_classes[-1]  # this will extract FYP prefer list
    matched = False
    # print("preferred list:",preferedlist)
    # print("FYP list", FYPlist)
    for prefered in preferedlist:
        if isinstance(prefered, list) and not matched:
            val = is_fit_maj(prefered, input_class, major_dict)
            if val != False:
                score = score_list[i] + val
                matched = True
        elif len(prefered) == 3 and not matched:
            if is_fit_field(prefered, input_class):
                score = score + score_list[i]
                matched = True
        elif input_class[:3] == "FYI" and not matched:
            val = is_fit_FYI(FYPlist, input_class)
            if val != False:
                score = score + val
                matched = True
        elif not matched:
            if is_fit_course(prefered, input_class):
                score = score + score_list[i]
                matched = True
        i += 1
    return score


# Tuan
# Added script to use score assignment 2
# Create manager
def main():
    manager = ClassAssignment.create_manager()
    print('Data import complete')
    print()

    students = manager.get_students()
    students.pop()

    classes = manager.get_courses()
    major_dict = map_courses_to_major(students, classes)

    default_dict = {}
    for a_student in students:
        student_ID = a_student.getID()
        student_preferred_courses = a_student.get_preferred_courses()
        for a_class in classes:
            class_name = a_class.getName()
            score = score_assignment_2(student_preferred_courses, class_name, major_dict)
            # print(student_ID + '//' + class_name + ":", score)
            # print('-------------------------------------------------')
            default_dict[(student_ID, class_name)] = score

    print('Fast tests for scoring function 2')

    print('Scoring student 2558905 with AST-051-01 Intro to Astronomy || Expected: 5, Real:',
          default_dict[('2558905', '*AST-051-01 Intro to Astronomy')])

    print('Scoring student 2558905 with AST-051L-01 Intro to Astronomy Lab || Expected: 0, Real:',
          default_dict[('2558905', 'AST-051L-01 Intro to Astronomy Lab')])

    print('Scoring student 2558905 with AST-051L-02 Intro to Astronomy Lab || Expected: 0, Real:',
          default_dict[('2558905', 'AST-051L-02 Intro to Astronomy Lab')])

    print('Scoring student 2113211 with PHL-100-01 Intro to Philosophy || Expected: 7.4, Real:',
          default_dict[('2113211', 'PHL-100-01 Intro to Philosophy')])

    print('Scoring student 2113211 with RUS-100-01 Basic Russian 1 || Expected: 7.3, Real:',
          default_dict[('2113211', 'RUS-100-01 Basic Russian 1')])

    print('Scoring student 2558459 with FYI-100-15 Music & Politics || Expected: 3, Real:',
          default_dict[('2558459', 'FYI-100-15 Music & Politics')])

    print('Scoring student 2558459 with FYI-100-16 Music & Politics || Expected: 3, Real:',
          default_dict[('2558459', 'FYI-100-16 Music & Politics')])

    print('----------------------------------------------------------------')


if __name__ == '__main__':
    main()
