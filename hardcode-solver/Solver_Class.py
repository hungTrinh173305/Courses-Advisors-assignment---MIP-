import sys
sys.path.append("../")

import time
import gurobipy as gp
from gurobipy import GRB
import copy
from backend.MappingMajorsToCourses import map_courses_to_major
import backend.ClassAssignment as ClassAssignment
import backend.ExcelManagement as em
from backend.Student import Student
from backend.Course import Course
import backend.class_scoring_v2_0 as scoring


HIGH_PEN = 10000
MED_PEN = 100
LOW_PEN = 1


def optimize_relaxation(m, m2, apply_to_m):
    '''Takes an original Gurobi model m, a copy of m named m2 that has had
    slack variables introduce to determine which constraints of m need
    to be relaxed, and Boolean flag apply_to_m.  This function
    attempts to solve m2 for the appropriate constraints of m to
    relax.  If m is already feasible, this function should return an
    empty list (and is also a waste of time to run).

    If m2 is not infeasible, which means that even under relaxation,
    there is still no solution to m, this function reports an error.  Then,
    computes IIS on m2, saving it in model_iis.ilp, and exits the
    program.

    If m2 is feasible, which means that there is a relaxation of some
    of the constraints of m that has a solution.  This function
    returns a list of the constraints that need to be relaxed, paired
    with how much to relax the RHS of the constraint by.  If
    apply_to_m is True, the appropriate constraints of m are relaxed
    in the way calculated.

    '''
    
    # This is relaxation is somewhat imprecise anyway because we can't
    # feasbility simulatenously optimize for the original objective,
    # so don't force it to calculate a fully optimal relaxation.
    # Also, the relaxation has a quadratic objective so it is much
    # harder for Gurobi to compute.
    m2.setParam("MIPGap", 0.7)  
    m2.update()
    # Solve for the relaxation of m.
    m2.optimize()

    status = m2.getAttr("Status")
    if status != GRB.OPTIMAL:
        print("+ Fatal Error: Mofel m is infeasible even when all soft constraints are relaxed!")
        print("+ Determining inconsistent hard constraints and outputing them in model_iis.ilp...")
        m2.computeIIS()
        m2.write("model_iis.ilp")  # Slow, but helpful to determine constraint violations.
        # The hard constraints are inconsistent, there is no solution
        # possible.  No choice but to give up and hope someone can use
        # model_iis.ilp to correct the hard constraints or turn some
        # into soft ones.
        exit(-3)

    # Approximately optimal relaxation has been computed, collect what
    # constraints need to be relaxed in m, updating m is apply_to_m is
    # set.
    relax_list = []  # Collects constraints to relax.
    for v in m2.getVars():
        # feasRelax injects m2 with new variables to allow slack in constraints.
        # The constraint named "my_constr" in m:
        #
        # x + y < 5
        #
        # Would be relaxed to in m2:
        #
        # x + y < 5 + ArtN_my_constr
        #
        # Where ArtN_my_constr is a new "slack" variable that takes a
        # positive value.  As Gurobi attempts to solve m2, it tries to
        # minimize the value of the added slack variables (so as to
        # relax the constraints as little as possible). So when the
        # optimization of m2 is complete.  The values of variable that
        # begin with ArtP_my_constr contain out much to relax the
        # constraint my_constr by.  Say ArtN_my_constr is assigned 3,
        # then the relaxed constraint in m should be:
        #
        # x + y < 5 + 3
        #
        # ArtN_ behaves similarly, but for constraints where slack
        # would need to substract from the RHS, e.g.,
        #
        # u + v > 7, might become u + v > 7 - ArtP_my_constr2
        
        if (v.VarName.startswith('ArtN_') or v.VarName.startswith('ArtP_')) and v.X > 1e-8:
            # https://support.gurobi.com/hc/en-us/community/posts/360074265771-Constraint-Relaxation
            is_N = v.VarName.startswith('ArtN_')
            cname = v.VarName[5:]
            con = m.getConstrByName(cname)
            if not con:
                # As an undocumented feature of the feasRelax, it
                # stops actually using string names for the new slack
                # variables and starts using their integer index.
                cnum = int(cname[1:])
                con = m.getConstrs()[cnum]
                
            # By this point con should be a constraint in m that needs
            # to relaxed.

            # Determine what kind of constraint con is.
            sense = con.getAttr(GRB.Attr.Sense)
            if sense == "<":
                # print(f"Relaxing {v.VarName[5:]} by {round(v.X)}")
                relax_list.append((con, round(v.X)))
                assert (is_N)
                if apply_to_m:
                    con.rhs += round(v.X)
            elif sense == ">":
                # print(f"Relaxing {v.VarName[5:]} by {round(-v.X)}")
                relax_list.append((con, round(-v.X)))
                assert (not is_N)
                if apply_to_m:
                    con.rhs += round(-v.X)
            elif sense == "=":
                if is_N:
                    # print(f"Relaxing {v.VarName[5:]} by {round(v.X)}")
                    relax_list.append((con, round(v.X)))
                    if apply_to_m:
                        con.rhs += round(v.X)
                else:
                    # print(f"Relaxing {v.VarName[5:]} by {round(-v.X)}")
                    relax_list.append((con, round(-v.X)))
                    if apply_to_m:
                        con.rhs += round(-v.X)

    return relax_list


def main():
    start_time = time.perf_counter()
    # Create manager
    print("\nImporting data...")

    # Hack to suppress openpyxlwarning message.
    # For some reason this doesn't work in ExcelManagement.py...
    # import os, sys
    # devnull = open('/dev/null', 'w')
    # oldstdout_fno = os.dup(sys.stderr.fileno())
    # os.dup2(devnull.fileno(), 2)
    manager = ClassAssignment.create_manager()
    # os.dup2(oldstdout_fno, 2)
    major_dict = map_courses_to_major(manager.get_students(), manager.get_courses())

    # Process student and course data
    print("\nMarshaling data... ")
    students = {}
    courses = {}

    for student in manager.get_students():
        if student.getID():
            students[student.getID()] = student

    course_prefs = {}
    for course in manager.get_courses():
        courses[course.getName()] = course
        course_prefs[course.getName()] = 0
    print("+ Number of students: ", len(students))
    print("+ Number of courses:  ", len(courses))

    
    
    print("\nScoring student preferences... ")
    default_dict = {}
    total_original_score = 0
    for student in students:
        s = students[student]
        student_data = s.get_preferred_courses()

        datafile_courses = []
        for i in range(1, 8):
            datafile_courses.append(s.get_student_detail("Course " + str(i)))

        student_score = 0
        for course in datafile_courses:
            if course and course != "":
                student_score += scoring.score_assignment([student_data], [[student, course]], major_dict)[0][1]
        total_original_score += student_score
        
        #print(student, student_data, student_score)
        for course in courses:
            # MWA - Bit of a hack to use Song's scoring.
            # MWA - Should take the following info into account, but it probably doesn't.
            # 5. Students should be assigned to an intro-course in their
            #   declared major or highest major interest.
            # 6. Students should be assigned to their preferred courses.
            # 14. FYI classes should be assigned based on student preferences.
            # 15. FYI classes should be gender balanced to an equal ratio.

            score = scoring.score_assignment([student_data], [[student, course]], major_dict)[0][1]
            if score > 0:
                #print(course, score)
                course_prefs[course] += score
            default_dict[(student, course)] = score

    print("+ Total score of original assignments:", total_original_score)
            
    # Check whether preferences have considered most classes.
    displayed_no_prefs = True
    for course in courses:
        if (course_prefs[course] == 0 and not courses[course].isLab()):
            if displayed_no_prefs:
                print(("+ Warning: Some courses have no students that prefer them, perhaps"
                       " the course are named / numbered inconsistently:"))
                displayed_no_prefs = False
            print(course)
    
    print("\nChecking scheduling conflicts...")
    combinations, scores = gp.multidict(default_dict)
    # print(scores)

    num_conflicts = 0
    for course1 in courses:
        for course2 in courses:
            if (course1 != course2 and courses[course1].hasOverlap(courses[course2])):
                # print("Timing conflict between", course1, course2)
                num_conflicts += 1

    print("+ Total time conflicts:", num_conflicts)

    print("\nCreating Gurobi Model...")
    # Create model
    print("+ Suppressing Gurobi Output!")
    env = gp.Env(empty=True)
    env.setParam("OutputFlag", 0)
    env.start()
    m = gp.Model("ClassAssignment", env=env)
    m.setParam(GRB.Param.DualReductions, 0)
    x = m.addVars(combinations, vtype=GRB.BINARY, name="assign")
    # Create constraints
    print("\nCreating Policy Constraints...")
    soft_constraints = []
    soft_weights = []
    manual_soft_constraints = []
    manual_soft_weights = []

    # 1. Each student is assigned to three classes excluding labs.
    print("+ HARD - Requiring students to take three lecture courses (except SCH-001)...")
    three_course_cons = m.addConstrs((gp.quicksum(x[s, name]
                                                  for (name, course) in courses.items()
                                                  if (not course.isLab() and name[:7] != "SCH-001"))
                                      == 3
                                      for s in students),
                                     name="three_non-lab_courses_per_student")

    # 2. Each student assigned to a class with a lab is assigned to a corresponding lab section.
    print("+ HARD - Requiring students to take a lab of courses they're in with labs...")
    lab_cons = m.addConstrs((
        (gp.quicksum(x[s, namemain]
                     for (namemain, coursemain) in courses.items()
                     if (coursemain.getDepartment() == course.getDepartment()
                         and coursemain.getNumber() == course.getNumber()))
         == gp.quicksum(x[s, namelab]
                        for (namelab, courselab) in courses.items()
                        if courselab.isLabOf(course)))
        for s in students
        for (name, course) in courses.items() if course.hasLab()),
        name="must_take_lab_section_with_lab_course")
    # 3. No student's class or lab assignments can overlap in timing.
    print("+ HARD - Preventing students from taking courses that overlap...")
    no_conflict_cons = m.addConstrs(
        (x[s, name1] + x[s, name2] <= 1
         for s in students
         for (name1, course1) in courses.items()
         for (name2, course2) in courses.items()
         if (name1 != name2 and course1.hasOverlap(course2))),
        name="no_time_conflicts_between_courses")

    # 4. No student will be assigned to more than 1 class in a single department.
    print("+ HARD - Preventing students from taking multiple lecture in same department....")
    no_duplicate_department_cons = m.addConstrs(
        (x[s, name1] + x[s, name2] <= 1
         for s in students
         for (name1, course1) in courses.items()
         for (name2, course2) in courses.items()
         if (name1 != name2
             and (not course1.isLab())
             and (not course2.isLab())
             and (course1.getDepartment() == course2.getDepartment()))),
        name="no_duplicate_department_among_courses")

    # 7. No course section can have more students assigned to it than the number of seats available.

    # MWA - Some courses have negative seats to begin with, e.g., MLT-200.
    # MWA - This is being violated by original data set.
    print("+ HARD (Manual Changes) - Preventing courses from enrolling more students than seats available...")
    course_seats_cons = m.addConstrs(
        (gp.quicksum(x[s, name] for s in students) <= max(course.getOpenSlots(), 0)
         for (name, course) in courses.items()),
        name="no_more_students_than_seats_in_section")

    manual_soft_constraints += [c for (n, c) in course_seats_cons.items()]
    manual_soft_weights += [0.2] * len(course_seats_cons)

    # 8. All course sections should have at least 6 students enrolled.

    # MWA - This being violated by the original dataset.
    print("+ SOFT (Low penalty) - Preventing courses from enrolling few total students than 6...")
    course_not_empty_cons = m.addConstrs(
        (gp.quicksum(x[s, name]
                     for s in students)
         >= 6 - max((course.getMaxStudents() - course.getOpenSlots()), 0)
         for (name, course) in courses.items()),
        name="courses_should_have_at_least_6_students.")

    soft_constraints += [c for (n, c) in course_not_empty_cons.items()]
    soft_weights += [LOW_PEN] * len(course_not_empty_cons)

    # 9. Students in certain circumstances will be required to take
    # certain courses: Declared or first major interest Engineering
    # students are assigned to ESC-100; LIM students are assigned to
    # HCX-630-51; Scholars are assigned to SCH-001 in ADDITION to 3
    # classes.

    print("+ HARD - Requiring LIM students to take HCX...")
    lim_take_hcx_cons = m.addConstrs(
        ((x[s, name] == (1 if ("LIM" == students[s].get_student_detail("HP/AP/US/ST/LIM/ Hockey")) else 0))
         for s in students
         for (name, course) in courses.items()
         if (name[:3] == "HCX")),
        name="only_LIM_must_take_HCX")

    print("+ SOFT (High penalty) - Requiring Scholar students to take SCH...")
    scholars_take_sch_cons = m.addConstrs(
        (x[s, name] == (1 if ("SCHOLARS" == students[s].get_student_detail("Assign Fall FYI")) else 0)
         for s in students
         for (name, course) in courses.items()
         if (name[:3] == "SCH")),
        name="only_scholars_must_take_SCH")

    soft_constraints += [c for (n, c) in scholars_take_sch_cons.items()]
    soft_weights += [HIGH_PEN] * len(scholars_take_sch_cons)

    print("+ HARD - Requiring engineering students to take ESC...")
    engineers_take_esc_cons = m.addConstrs(
        (gp.quicksum(x[s, name]
                     for (name, course) in courses.items()
                     if name[:4] == "*ESC")
         == 1
         for s in students
         if (students[s].get_student_detail("Engineering?") == "Yes")),
        name="engineers_must_take_ESC")

    # 10. NOT IMPLEMENTED - If a student has taken an AP exam to test
    # out of a class, they must not be assigned to that class.

    # MWA - This is repeatedly violated in the dataset.  AP exams
    # taken seem to be ignored by the dataset, e.g., many students
    # have taken AP chem, but are placed into section by department or
    # dean.  Maybe using outside info, need more robust explanation to
    # actually implement.
    # 11. If Paul has created a math placement for a student, that
    # student must be assigned to a section of that math course.

    # MWA - Data should be santized by hand in the future.
    # MWA - Don't assign anyone not marked as fall else?
    # MWA - The MTH-113, MTH-115H section size is violated by the data set.

    print("+ HARD - Requiring students to take recommended MTH courses...")
    take_mth_cons = m.addConstrs(
        ((gp.quicksum(x[s, name]
                     for (name, course) in courses.items()
                     if students[s].get_student_detail("MTH") in name))
         == 1
         for s in students
         if students[s].get_student_detail("MTH") is not None),
        name="take_recommended_math")

    print("+ HARD - Requiring students not recommended for MTH courses to take no MTH courses...")
    take_no_mth_cons = m.addConstrs(
        ((gp.quicksum(x[s, name]
                     for (name, course) in courses.items()
                     if name[:3] == "MTH"))
        == 0
         for s in students
         if students[s].get_student_detail("MTH") is None),
        name="no_math_if_not_recommended")

    # 12. If Daniel has created a language placement for a student,
    # that student must be assigned to a section of that language
    # course

    # MWA - Data should be santized by hand in the future.
    # MWA - There are students in the dataset that must take HCX, MTH, CHM, LANG, which is infeasible.
    # MWA - SPN-100 has -1 seats to begin with dataset suggests 19 students take SPN 100...

    print("+ SOFT (Low Penalty) - Requiring students to take recommended language courses...")
    take_lang_cons = m.addConstrs(
        ((gp.quicksum(x[s, name]
                     for (name, course) in courses.items()
                     if (students[s].get_student_detail("Lang")
                         in name)))
         == 1
         for s in students
         if students[s].get_student_detail("Lang") is not None),
        name="take_recommended_lang")

    soft_constraints += [c for (n, c) in take_lang_cons.items()]
    soft_weights += [LOW_PEN] * len(take_lang_cons)

    print("+ HARD - Requiring students not recommended for language courses to take no language courses...")
    lang_codes = ["LAT", "SPN", "CHN", "FRN", "JPN", "GER", "ARB", "HEB", "GRK", "RUS"]

    dont_take_lang_cons = m.addConstrs(
        ((x[s, name] == 0)
         for s in students
         for (name, course) in courses.items()
         if (name[:3] in lang_codes
             and students[s].get_student_detail("Lang") is None)),
        name="dont_take_lang")

    # 13. If the chemistry department has created a chemistry
    # placement for a student, that student must be assigned to a
    # section of that chemistry course.
    # MWA - Missing condition to exclude those not listed for CHM-110H
    # MWA - The CHM-110H section size is violated by the data set, should force class limits up.

    print("+ SOFT (Low Penalty) - Requiring students to take recommended CHM courses...")
    take_chm_cons = m.addConstrs(
        ((gp.quicksum(x[s, name]
                     for (name, course) in courses.items()
                     if students[s].get_student_detail("CHM") in name))
         == 1
         for s in students
         if students[s].get_student_detail("CHM") is not None),
        name="take_recommended_chm")

    soft_constraints += [c for (n, c) in take_chm_cons.items()]
    soft_weights += [LOW_PEN] * len(take_chm_cons)

    print("+ HARD - Requiring students not recommended for CHM-110H to not take it...")
    dont_take_honors_chm_cons = m.addConstrs(
        ((x[s, name] == 0)
         for s in students
         for (name, course) in courses.items()
         if (name[:8] == "CHM-110H"
             and students[s].get_student_detail("CHM") is None)),
        name="dont_take_honors_chm_if_not_recommended")

    # 15. FYI classes should be gender balanced to an equal ratio.
    print("+ SOFT (Low Penalty) - FYI classes should be gender balanced to be equal ratio...")

    # MWA - The incoming class gender balance is substantally off
    # over, this will make this impossible, or make the gender balance
    # of winter FYI's even more off.

    fyi_gender_balanced_cons = m.addConstrs(
        ((gp.quicksum(x[s, name]
                      for s in students
                      if students[s].get_student_detail("Gender") == "F")
          == (course.getOpenSlots() // 2)
          for (name, course) in courses.items() if name[:3] == "FYI")),
        name="fyi_gender_balanced")

    soft_constraints += [c for (n, c) in fyi_gender_balanced_cons.items()]
    soft_weights += [LOW_PEN] * len(fyi_gender_balanced_cons)

    # 16. All FYI slots must be filled
    print("+ HARD - Requiring all FYI seats be filled...")
    fyi_filled_cons = m.addConstrs(
        (gp.quicksum(x[s, name] for s in students) == max(course.getOpenSlots(), 0)
         for (name, course) in courses.items() if name[:3] == "FYI"),
        name="all_fyi_seats_filled")

    # 17. Scholars and LIM may not be assigned to a Fall FYI
    print("+ HARD - Preventing LIM and scholars from taking FYI...")
    no_scholars_fyi_cons = m.addConstrs(
        ((x[s, name] == 0)
         for s in students
         for (name, course) in courses.items()
         if (name[:3] == "FYI"
             and (students[s].get_student_detail("HP/AP/US/ST/LIM/ Hockey") in ["LIM", "US"]))),
        name="no_scholars_in_fyi")


    
    # Policy has been expressed by constraints, now optimize!
    print("\nSolving for optimal solution...")

    m.write("hard_class_model_init.mps")
    m.setObjective(x.prod(scores), GRB.MAXIMIZE)
    m.optimize()

    # Deal with violated constraints.
    status = m.getAttr("Status")
    while status != GRB.OPTIMAL:

        # Error States
        

        if status == GRB.INF_OR_UNBD:
            # If this is the case, something is really broken, it
            # should never occur with the current settings.
            print("+ Fatal Error: Solution is infeasible or objective is unbounded!")
            exit(-2)

        elif status == GRB.INFEASIBLE:
            # This occurs when there is no solution that satisfies all
            # the constraints.
            print("+ Error: Constraints are inconsistent, no solution found.  Some constraints must be relaxed.")

            # If there are no soft constraints there's nothing to do
            # by give up.  Probably some hard constraints should be soft.
            if soft_constraints == [] and manual_soft_constraints == []:
                print("+ Fatal Error: Constraints violated, but none are soft.")
                exit(-1)

            m.reset()

            print("+ Suggesting manual adjustments to hard constraints...")
            # Make new model m2 that allows relaxation of some constraints.
            m2 = m.copy()
            m2.feasRelax(0, False, None, None, None,
                         soft_constraints + manual_soft_constraints,
                         soft_weights + manual_soft_weights)

            m2.write("hard_class_model_hard_relax.mps")

            
            suggested_relaxation = optimize_relaxation(m, m2, False)
            # Output only manual relaxation suggestions in client
            # readable text.  The client will have to modify the
            # dataset to resolve these conflicts, probably after
            # consultation with a third-party.  We never automatically
            # change m to relax the manual soft constraints.
            for con1 in manual_soft_constraints:
                acc = 0
                for (con2, v) in suggested_relaxation:
                    if con1 == con2:
                        acc += v
                if round(acc) != 0:
                    # This is specific to the types of soft manual
                    # constraints included, update this if those
                    # change.  Now it's only course capacity constraints.
                    course_name = str(con1).split("[")[1].split("]")[0]
                    print(f"-- Increase max seats in {course_name} by {acc}.")

            print("+ Relaxing soft constraints...")
            # Make new model m2 that allows relaxation of only the
            # soft constraints, but not the soft manual constraints.
            m2 = m.copy()
            m2.feasRelax(0, False, None, None, None, soft_constraints, soft_weights)
            m2.write("hard_class_model_soft_relax.mps")
            
            suggested_relaxation = optimize_relaxation(m, m2, True)

            # Output relaxations done, converting to client readable
            # text.
            for con1 in soft_constraints:
                acc = 0
                for (con2, v) in suggested_relaxation:
                    if con1 == con2:
                        acc += v
                if round(acc) != 0:
                    # This is specific to the types of soft
                    # constraints include, update this if those
                    # change.
                    if "only_scholars_must_take_SCH" in str(con1):
                        student = str(con1).split("[")[1][:7]
                        print(f"-- Permit student {student} not to take SCH-001.")
                    elif "take_recommended_lang" in str(con1):
                        student = str(con1).split("[")[1][:7]
                        print(f"-- Permit student {student} not to take recommended language course.")
                    elif "take_recommended_chm" in str(con1):
                        student = str(con1).split("[")[1][:7]
                        print(f"-- Permit student {student} not to take recommended CHM course.")
                    else:
                        print(f"-- Relax {con1} by {acc}.")
            print("+ Solving with relaxed soft constraints...")

            # Resolve m using the relaxed soft constraints.
            m.update()
            m.reset()
            m.write("hard_class_model_final.mps")
            m.optimize()
            # Optimization should always succeed here, because we
            # sufficiently relaxed constraints that were causing
            # infeasibility.
            
        elif status == GRB.UNBOUNDED:
            # This shouldn't ever happen for real data, if it does it
            # indicate a programmer error.
            print("+ Fatal Error: Objective is unbounded, check that all variables are constrained.")
            exit(-1)
        status = m.getAttr("Status")

    print("\nSolving Complete!")
    print('+ Total score - our solution:     ', m.objVal)
    print('+ Total score - original dataset: ', total_original_score)    
    print(f'+ Score improvement: {(m.objval - total_original_score) / total_original_score * 100:.2f}%')
    
    outfile_name = "../output_files/ClassOutput.xlsx"
    print(f"\nOutputting Solution to {outfile_name}...")

    # Clear sheet.
    wipe_data = [[""] * 8] * (len(students) + 1)
    em.write_workbook("Class-Assignment", wipe_data, outfile_name)

    # Output contents into sheet.
    headers = ["Colleague ID Anonymized", "Course 1", "Course 2", "Course 3",
               "Course 4", "Course 5", "Course 6", "Course 7"]
    out_data = [headers]
    for student in students.keys():
        row = [student]
        for course in courses.keys():
            if x[student, course].x > 1e-6:
                row.append(course)
        out_data.append(row)

    em.write_workbook("Class-Assignment", out_data, outfile_name)

    print("\nSolving Complete...")
    
    print(f"+ Runtime {time.perf_counter() - start_time:.2f} seconds.")


if __name__ == '__main__':
    main()
