# SolverFinish Function

## Business Description
Tells Microsoft Office Excel what to do with the results and what kind of report to create when the solution process is completed.

## Behavior
Tells Microsoft Office Excel what to do with the results and what kind of report to create when the solution process is completed.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
SolverLoad LoadArea:=Range("A33:A38") 
SolverSolve UserFinish:=TrueSolverFinishKeepFinal:=1, ReportArray:=Array(1)
```