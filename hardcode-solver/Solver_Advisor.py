import sys
sys.path.append("../")

import gurobipy as gp
from gurobipy import GRB

import backend.AdvisorAssignment as AdvisorAssignment
import backend.ExcelManagement as em
import backend.Student as Student
import backend.Advisor as Advisor

OUTFILE = "../output_files/ClassOutput.xlsx"
OUTTAB = "Advisor-Assignment"

def main():
    # Create manager
    
    manager = AdvisorAssignment.create_manager()
    
    print("Data import complete")
    
    # Create model
    
    student_IDs = []
    advisor_names = []
    
    students = manager.get_students()
    advisors = manager.get_advisors()
    
    for student in students:
        student_IDs.append(student.getID())
        
    for advisor in advisors:
        advisor_names.append(advisor.getName())
        
    default_dict = {}
    for student in student_IDs:
        for advisor in advisor_names:
            student_obj = manager.get_student(student)
            advisor_obj = manager.get_advisor(advisor)
            scoring = manager.score_assignment(student_obj, advisor_obj)
            # if scoring > 2:
            #     print(scoring)
            
            default_dict[(student, advisor)] = scoring
    
    combinations, scores = gp.multidict(default_dict)
    
    m = gp.Model("AdvisorAssignment")
    x = m.addVars(combinations, name="assign")
    
    # Create constraints
    
    LIM_advisor = "Cuddy, L."
    
    student_constraints = m.addConstrs((x.sum(s, "*") == 1 for s in student_IDs), name="one advisor per student")
    
    advisor_names.remove(LIM_advisor) #LIM advisor can get more than 8 students
    #commented out because it is really slow. TODO: Change AdvisorAssignment to use dictionary instead of list
    #advisor_constraints = m.addConstrs((x.sum("*", a) <= manager.get_advisor(a).getMaxStudents() for a in advisor_names), name="fewer than 8 students per advisor")
    advisor_constraints = m.addConstrs((x.sum("*", a) <= 8 for a in advisor_names), name="fewer than 8 students per advisor")
    
    advisor_names.append(LIM_advisor)
    
    
    #All and only LIM students assigned to LIM advisor
    
    for student in manager.get_students():
        ID = student.getID()
        if student.get_student_detail("HP/AP/US/ST/LIM/ Hockey") == "LIM":
            m.addConstr(x[(ID, LIM_advisor)] == 1, name="LIM Constraints")
        else:
            m.addConstr(x[(ID, LIM_advisor)] == 0, name="LIM Constraints")

    m.setObjective(x.prod(scores), GRB.MAXIMIZE)
    m.write("model.mps")
    m.optimize()

    # for v in m.getVars():
    #     if v.x > 1e-6:
    #         print(v.varName, v.x)
    # print('Total matching score: ', m.objVal)
    
    # make pairings in manager
    
    counter = 1
    col_1 = ["Student_ID"]
    col_2 = ["Calculated advisor"]
    col_3 = ["Calculated score"]
    for s, a in combinations:
        if x[s, a].x > 1e-6:
            score = default_dict[(s, a)]
            col_1.append(s)
            col_2.append(a)
            col_3.append(score)
            # em.write_workbook_cell("Assignments", counter, 0, s, filename=OUTFILE)
            # em.write_workbook_cell("Assignments", counter, 1, a, filename=OUTFILE)
            # score = default_dict[(s, a)]
            # em.write_workbook_cell("Assignments", counter, 2, score, filename=OUTFILE)
            counter += 1
    em.write_workbook_col(OUTTAB, 0, col_1, filename=OUTFILE)
    em.write_workbook_col(OUTTAB, 1, col_2, filename=OUTFILE)
    em.write_workbook_col(OUTTAB, 2, col_3, filename=OUTFILE)


if __name__ == '__main__':
    main()
    
