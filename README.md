
## Project Overview
- At Union College, the Register office has to assign incoming freshmen into their preferred courses and appropriate academic advisors annually. Each student was asked to fill out a form that would provide their rankings of preferences in courses and majors. Based on that rankings, the task is to assign students into courses and advisors so that everyone should be happy with the assignments.
The job involves about 600 students, 200 courses, and 160 professors. It usually takes the staffs  3 months to accomplish the job by hands.  
- Our solution treated the assignments as a mixed integer programming problem. Using Gurobi model to maximize the overall score of every assignment, we are
able to reach a solution not only in a couple of minutes, but also provides better results, comparing to human assignments.

## Installation & Dependencies
Requires Python 3 with numpy, gurobipy, panda, openpyxl, pydot

## Organization
1. ```input_files/``` - Contains input xlsx data.
2. ```output_files/``` - Contains output xlsx data and analysis files.
3. ```hardcode-solver/``` - Contains hardcoded solvers for advisor and class assignment and corresponding analysis programs. Run solver (e.g., ```python3 Solver_Class.py```), then scoring (e.g., ```python3 Scoring_Class.py```).  Output is in ```output_files/ClassOutput.xlsx```.
4. ```backend/``` - Contains common data structures and testing code.


