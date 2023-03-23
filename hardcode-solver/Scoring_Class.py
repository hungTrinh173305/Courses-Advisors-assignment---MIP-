######### Only run after Solver_Class is finished! ##########

import sys
sys.path.append("../")
import backend.ExcelManagement as em
from backend.Student import Student
from pprint import pprint
import backend.class_scoring_v2_0 as scoring
from backend.MappingMajorsToCourses import map_courses_to_major
import backend.ClassAssignment as ClassAssignment


DEFAULT_INPUT = "../input_files/Copy of Class of 2026 (Fall 2022) Advisor and Registration Assignments - Anonymized.xlsx"
DEFAULT_OUTPUT = "../output_files/ClassOutput.xlsx"

def is_clean(list):
    # if there is NONE in list, return False otherwise return True
    if None in list:
        return False
    else:
        return True


def remove_NONE(list):
    # remove all NONE in a list
    clear = is_clean(list)
    while not clear:
        list.remove(None)
        clear = is_clean(list)


def remove_irr(course):
    # format assigned class
    # print(course)
    if course != None:
        if course[0] == "*":
            course = course[1:]
        if course[7] == "L":
            course = None
    return course


def mapping_com(worksheet="Class-Assignment", workbook=DEFAULT_OUTPUT):
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
    for i in range(1, 580):  # skip first row - header row
        to_append = []
        n = 0
        for cell in major_to_dept_spreadsheet[i]:
            if n == 0:
                to_append.append(cell)
            elif cell is not None and n < 8:
                to_append.append(remove_irr(cell))
            n += 1
        remove_NONE(to_append)
        mapping.append(to_append)
    return mapping


def mapping_hum(worksheet="Main Tab",
                workbook=DEFAULT_INPUT):
    """
    creates a mapping between student majors and advisor departments based on an Excel workbook with the relevant
    information.
    @param worksheet: the name of the worksheet this information is stored in. defaults to Sheet1, the default
    worksheet name in Excel.
    @param workbook: the name of the workbook (i.e. Excel file) this information is stored in. Defaults to
    major_to_dept_matching.xlsx, the file in which the data is stored.
    """
    # TT: should use the get_student_detail('Course 1') function in Student.py, instead of reading the spreadsheet
    # again, which would take much more time to run.

    mapping = []
    major_to_dept_spreadsheet = em.read_workbook(worksheet, workbook)
    for i in range(1, 580):  # skip first row - header row
        to_append = []
        row = major_to_dept_spreadsheet[i]
        to_append.append(row[0])
        to_append.append(remove_irr(row[79]))
        to_append.append(remove_irr(row[90]))
        to_append.append(remove_irr(row[101]))
        to_append.append(remove_irr(row[112]))
        to_append.append(remove_irr(row[123]))
        remove_NONE(to_append)
        mapping.append(to_append)
    return mapping


def counting(input_class, score, counting):
    # take the prefered class list and assigned class list
    # return a data list including ID, match score and number of class
    # TT: should use student.get_student_detail('Preferred Course_1') in the conditional statements instead of using the scoring metrics.
    data = []
    if input_class[:3] != "FYI":
        if score == 49:
            counting[0] = counting[0] + 1
        elif score == 36:
            counting[1] = counting[1] + 1
        elif score == 25:
            counting[2] = counting[2] + 1
        elif score == 16:
            counting[3] = counting[3] + 1
        elif score == 9:
            counting[4] = counting[4] + 1
        elif score == 4:
            counting[5] = counting[5] + 1
        elif score == 1:
            counting[6] = counting[6] + 1
        else:
            counting[7] = counting[7] + 1
    else:
        if score == 36:
            counting[8] = counting[8] + 1
        elif score == 25:
            counting[9] = counting[9] + 1
        elif score == 16:
            counting[10] = counting[10] + 1
        elif score == 9:
            counting[11] = counting[11] + 1
        elif score == 4:
            counting[12] = counting[12] + 1
        elif score == 1:
            counting[13] = counting[13] + 1
        else:
            counting[14] = counting[14] + 1
    return counting


def analyze_gender(map_hum, map_com, manager, major_dict):
    gender_dict = {'M by Com': 0, 'F by Com': 0, 'M by Human': 0, 'F by Human': 0}

    total_M_com = 0
    total_F_com = 0
    total_M_hum = 0
    total_F_hum = 0

    # scoring human assigned
    for i in range(len(map_hum)):
        studentID = map_hum[i][0]
        classes_assigned = map_hum[i][1:]
        student_obj = manager.get_student(studentID)
        gender = student_obj.get_student_detail('Gender')
        if gender == 'M':  # scoring Male students assigned by human
            for a_class_name in classes_assigned:
                pref_list = student_obj.get_preferred_courses()
                class_score = scoring.score_assignment_2(pref_list, a_class_name, major_dict)
                total_M_hum += class_score
        elif gender == 'F':  # scoring Male students assigned by human
            for a_class_name in classes_assigned:
                pref_list = student_obj.get_preferred_courses()
                class_score = scoring.score_assignment_2(pref_list, a_class_name, major_dict)
                total_F_hum += class_score

    # scoring com assigned
    for i in range(len(map_com)):
        studentID = map_com[i][0]
        classes_assigned = map_com[i][1:]
        student_obj = manager.get_student(studentID)
        gender = student_obj.get_student_detail('Gender')
        if gender == 'M':  # scoring Male students assigned by human
            for a_class_name in classes_assigned:
                pref_list = student_obj.get_preferred_courses()
                class_score = scoring.score_assignment_2(pref_list, a_class_name, major_dict)
                total_M_com += class_score
        elif gender == 'F':  # scoring Male students assigned by human
            for a_class_name in classes_assigned:
                pref_list = student_obj.get_preferred_courses()
                class_score = scoring.score_assignment_2(pref_list, a_class_name, major_dict)
                total_F_com += class_score

    gender_dict['F by Com'] = total_F_com
    gender_dict['F by Human'] = total_F_hum
    gender_dict['M by Com'] = total_M_com
    gender_dict['M by Human'] = total_M_hum

    return gender_dict


def main():
    print('Importing Data . . . ')
    manager = ClassAssignment.create_manager()
    print('Done import.')
    students = manager.get_students()
    students.pop()
    classes = manager.get_courses()
    map_com = mapping_com()
    map_hum = mapping_hum()
    major_dict = map_courses_to_major(students, classes)
    prefer = []
    hum_total = 0
    com_total = 0
    hum_rank = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    com_rank = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    for a_student in students:
        student_preferred_courses = a_student.get_preferred_courses()
        prefer.append(student_preferred_courses)

    hum_score = []
    if len(prefer) != len(map_hum):
        print(len(prefer), len(map_hum))
    for i in range(len(map_hum)):
        classlist = map_hum[i]
        preferedlist = prefer[i]
        classlist = classlist[1:]
        stu_score = 0
        for classes in classlist:
            score = scoring.score_assignment_2(preferedlist, classes, major_dict)
            stu_score = stu_score + score
            counting(classes, score, hum_rank)
        hum_total += stu_score
        hum_score.append(stu_score)

    com_score = []
    if len(prefer) != len(map_com):
        print(len(prefer), len(map_com))
    for i in range(len(map_com)):
        classlist = map_com[i]
        preferedlist = prefer[i]
        classlist = classlist[1:]
        stu_score = 0
        for classes in classlist:
            score = scoring.score_assignment_2(preferedlist, classes, major_dict)
            stu_score = stu_score + score
            counting(classes, score, com_rank)
        com_total += stu_score
        com_score.append(stu_score)

    input_file_name = DEFAULT_INPUT
    output_file_name = DEFAULT_OUTPUT
        
    write_list = ["Score of Computer Assigning"]
    for text in com_score:
        write_list.append(text)
    em.write_workbook_col("Class-Assignment", 8, write_list, filename=output_file_name)
    write_list = ["Score of Human Assigning"]
    for text in hum_score:
        write_list.append(text)

    print('Analyzing results . . .')
    print("Analyzing by students' preferences . . .")
    em.write_workbook_col("Class-Assignment", 16, write_list, filename=output_file_name)

    em.write_workbook_col("Class-Assignment", 9, em.read_workbook_col("Main Tab", 79,
                                                                      input_file_name),
                          filename=output_file_name)
    em.write_workbook_col("Class-Assignment", 10, em.read_workbook_col("Main Tab", 90,
                                                                       input_file_name),
                          filename=output_file_name)
    em.write_workbook_col("Class-Assignment", 11, em.read_workbook_col("Main Tab", 101,
                                                                       input_file_name),
                          filename=output_file_name)
    em.write_workbook_col("Class-Assignment", 12, em.read_workbook_col("Main Tab", 112,
                                                                       input_file_name),
                          filename=output_file_name)
    em.write_workbook_col("Class-Assignment", 13, em.read_workbook_col("Main Tab", 123,
                                                                       input_file_name),
                          filename=output_file_name)
    em.write_workbook_col("Class-Assignment", 14, em.read_workbook_col("Main Tab", 134,
                                                                       input_file_name),
                          filename=output_file_name)
    em.write_workbook_col("Class-Assignment", 15, em.read_workbook_col("Main Tab", 145,
                                                                       input_file_name),
                          filename=output_file_name)
    print("Analyzing total . . .")

    text = ["", "", "", "", "", "", "", "Total", com_total, "", "", "", "", "", "", "Total", hum_total]
    em.write_workbook_row("Class-Assignment", len(com_score) + 1, text, filename=output_file_name)

    print("Analyzing by gender . . .")
    a = analyze_gender(map_hum, map_com, manager, major_dict)

    text = ["", "", "", "", "", "", "", "Male", a['M by Com'], "", "", "", "", "", "", "Male", a['M by Human']]
    em.write_workbook_row("Class-Assignment", len(com_score) + 2, text, filename=output_file_name)

    text = ["", "", "", "", "", "", "", "Female", a['F by Com'], "", "", "", "", "", "", "Female", a['F by Human']]
    em.write_workbook_row("Class-Assignment", len(com_score) + 3, text, filename=output_file_name)

    text = ["", "Preference 1 Matched", "Preference 2 Matched", "Preference 3 Matched", "Preference 4 Matched",
            "Preference 5 Matched", "Preference 6 Matched", "Preference 7 Matched", "Match Failed"]
    em.write_workbook_row("Class-Analytics", 0, text, filename=output_file_name)

    text = ["Computer Assignment"]
    i = 0
    for number in com_rank[:8]:
        text.append(number)
        i += 1
    em.write_workbook_row("Class-Analytics", 1, text, filename=output_file_name)

    text = ["Human Assignment"]
    i = 0
    for number in hum_rank[:8]:
        text.append(number)
    em.write_workbook_row("Class-Analytics", 2, text, filename=output_file_name)

    text = ["", "FYP 1 Matched", "FYP 2 Matched", "FYP 3 Matched", "FYP 4 Matched",
            "FYP 5 Matched", "FYP Reason & Passion Matched", "Match Failed"]
    em.write_workbook_row("Class-Analytics", 3, text, filename=output_file_name)

    text = ["Computer Assignment"]
    i = 0
    for number in com_rank[8:]:
        text.append(number)
        i += 1
    em.write_workbook_row("Class-Analytics", 4, text, filename=output_file_name)

    text = ["Human Assignment"]
    i = 0
    for number in hum_rank[8:]:
        text.append(number)
    em.write_workbook_row("Class-Analytics", 5, text, filename=output_file_name)

    print('Done')


if __name__ == '__main__':
    main()
