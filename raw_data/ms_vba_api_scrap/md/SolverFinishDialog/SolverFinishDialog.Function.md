# SolverFinishDialog Function

## Business Description
Tells Microsoft Office Excel what to do with the results and what kind of report to create when the solution process is completed. Equivalent to the SolverFinish function, but also displays the Solver Results dialog box after solving a problem.

## Behavior
Tells Microsoft Office Excel what to do with the results and what kind of report to create when the solution process is completed. Equivalent to theSolverFinishfunction, but also displays theSolver Resultsdialog box after solving a problem.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
SolverLoad loadArea:=Range("A33:A38") 
SolverSolve userFinish:=TrueSolverFinishDialogkeepFinal:=1, reportArray:=Array(1)
```