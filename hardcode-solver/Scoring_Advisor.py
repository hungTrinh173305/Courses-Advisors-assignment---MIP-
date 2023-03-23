######### Only run after Solver_Advisor is finished! ##########


if __name__ == '__main__':
    pass

import sys
sys.path.append("../")
import backend.AdvisorAssignment as AdvisorAssignment
import backend.ExcelManagement as em
import backend.Student as Student
import backend.Advisor as Advisor

OUTFILE = "../output_files/ClassOutput.xlsx"
OUTTAB_ASSIGN = "Advisor-Assignment"
OUTTAB_ANALYTICS = "Advisor-Analytics"

def convert(score):
    if score == 10:
        return "Meet Major"
    if score == 5:
        return "Meet interest_1"
    if score == 4:
        return "Meet interest_2"
    if score == 3:
        return "Meet interest_3"
    if score == 2:
        return "Meet interest_4"
    if score == 1:
        return "Meet falied"


def count(score, grade):
    if score == 10:
        grade[0] = grade[0] + 1
    if score == 5:
        grade[1] = grade[1] + 1
    if score == 4:
        grade[2] = grade[2] + 1
    if score == 3:
        grade[3] = grade[3] + 1
    if score == 2:
        grade[4] = grade[4] + 1
    if score == 1:
        grade[5] = grade[5] + 1
    return grade
  
  
def main():
    # Create manager
    
    manager = AdvisorAssignment.create_manager()
    # print("created manager")
    
    # Create model
    
    student_IDs = []
    advisor_names = []
    
    students = manager.get_students()
    advisors = manager.get_advisors()
    
    for student in students:
        student_IDs.append(student.getID())
    
    for student in students:
        advisor_names.append(student.get_student_detail("Advisor"))
    
    cur_score = []
    col_4 = ["Original advisor"]
    col_5 = ["Original scoring"]
    for i in range(len(student_IDs)):
        student = student_IDs[i]
        advisor = advisor_names[i]
        student_obj = manager.get_student(student)
        advisor_obj = manager.get_advisor(advisor)
        scoring = manager.score_assignment(student_obj, advisor_obj)
        # print(student, advisor, scoring)
        cur_score.append(scoring)
        # em.write_workbook_cell(OUTTAB_ASSIGN, i + 1, 4, advisor, filename=OUTFILE)
        # em.write_workbook_cell(OUTTAB_ASSIGN, i + 1, 5, scoring, filename=OUTFILE)
        col_4.append(advisor)
        col_5.append(scoring)
    em.write_workbook_col(OUTTAB_ASSIGN, 4, col_4, filename=OUTFILE)
    em.write_workbook_col(OUTTAB_ASSIGN, 5, col_5, filename=OUTFILE)
    # print("wrote original adv/score")

    con_score = em.read_workbook_col(OUTTAB_ASSIGN, 2, filename=OUTFILE)[1:]
    
    con_grade = [0, 0, 0, 0, 0, 0]
    i = 0
    col_3 = ["Calculated met requirement"]
    for element in con_score:
        text = convert(element)
        # em.write_workbook_cell(OUTTAB_ASSIGN, i + 1, 3, text, filename=OUTFILE)
        col_3.append(text)
        count(element, con_grade)
        i += 1
    em.write_workbook_col(OUTTAB_ASSIGN, 3, col_3, filename=OUTFILE)
    # print("wrote calculated grades")
    
    cur_grade = [0, 0, 0, 0, 0, 0]
    i = 0
    col_6 = ["Original met requirement"]
    for element in cur_score:
        text = convert(element)
        # em.write_workbook_cell(OUTTAB_ASSIGN, i + 1, 6, text, filename=OUTFILE)
        col_6.append(text)
        count(element, cur_grade)
        i += 1
    em.write_workbook_col(OUTTAB_ASSIGN, 6, col_6, filename=OUTFILE)
    # print("wrote original grades")
    
    # print("Grade of Computer assigning:")
    # print("Match major:" + str(con_grade[0]) + " Match 1st interest:" + str(con_grade[1]) + " Match 2nd interest:"
    #       + str(con_grade[2]) + " Match 3rd interest:" + str(con_grade[3]) + " Match 4th interest:" + str(con_grade[4])
    #       + " Match Failed:" + str(con_grade[5]))
    # print("===============")
    # print("===============")
    # print("Grade of Human assigning:")
    # print("Match major:" + str(cur_grade[0]) + " Match 1st interest:" + str(cur_grade[1]) + " Match 2nd interest:"
    #       + str(cur_grade[2]) + " Match 3rd interest:" + str(cur_grade[3]) + " Match 4th interest:" + str(cur_grade[4])
    #       + " Match Failed:" + str(cur_grade[5]))
    
    
    em.write_workbook_cell(OUTTAB_ANALYTICS, 2, 0, "Human Assignment", filename=OUTFILE)
    em.write_workbook_cell(OUTTAB_ANALYTICS, 1, 0, "Computer Assignment", filename=OUTFILE)
    
    em.write_workbook_cell(OUTTAB_ANALYTICS, 0, 1, "Majors matched", filename=OUTFILE)
    em.write_workbook_cell(OUTTAB_ANALYTICS, 1, 1, con_grade[0], filename=OUTFILE)
    em.write_workbook_cell(OUTTAB_ANALYTICS, 2, 1, cur_grade[0], filename=OUTFILE)
    
    em.write_workbook_cell(OUTTAB_ANALYTICS, 0, 2, "First Preference or better", filename=OUTFILE)
    em.write_workbook_cell(OUTTAB_ANALYTICS, 1, 2, sum(con_grade[0:2]), filename=OUTFILE)
    em.write_workbook_cell(OUTTAB_ANALYTICS, 2, 2, sum(cur_grade[0:2]), filename=OUTFILE)
    
    em.write_workbook_cell(OUTTAB_ANALYTICS, 0, 3, "Second Preference or better", filename=OUTFILE)
    em.write_workbook_cell(OUTTAB_ANALYTICS, 1, 3, sum(con_grade[0:3]), filename=OUTFILE)
    em.write_workbook_cell(OUTTAB_ANALYTICS, 2, 3, sum(cur_grade[0:3]), filename=OUTFILE)
    
    em.write_workbook_cell(OUTTAB_ANALYTICS, 0, 4, "Third Preference or better", filename=OUTFILE)
    em.write_workbook_cell(OUTTAB_ANALYTICS, 1, 4, sum(con_grade[0:4]), filename=OUTFILE)
    em.write_workbook_cell(OUTTAB_ANALYTICS, 2, 4, sum(cur_grade[0:4]), filename=OUTFILE)
    
    em.write_workbook_cell(OUTTAB_ANALYTICS, 0, 5, "Fourth Preference or better", filename=OUTFILE)
    em.write_workbook_cell(OUTTAB_ANALYTICS, 1, 5, sum(con_grade[0:5]), filename=OUTFILE)
    em.write_workbook_cell(OUTTAB_ANALYTICS, 2, 5, sum(cur_grade[0:5]), filename=OUTFILE)
    
    em.write_workbook_cell(OUTTAB_ANALYTICS, 0, 6, "Match Failed", filename=OUTFILE)
    em.write_workbook_cell(OUTTAB_ANALYTICS, 1, 6, con_grade[5], filename=OUTFILE)
    em.write_workbook_cell(OUTTAB_ANALYTICS, 2, 6, cur_grade[5], filename=OUTFILE)
    

if __name__ == '__main__':
    main()
