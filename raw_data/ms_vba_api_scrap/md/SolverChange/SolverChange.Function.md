# SolverChange Function

## Business Description
Changes an existing constraint. Equivalent to clicking Solver in the Data | Analysis group and then clicking Change in the Solver Parameters dialog box.

## Behavior
Changes an existing constraint. Equivalent to clickingSolverin theData|Analysisgroup and then clickingChangein theSolver Parametersdialog box.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
SolverLoad loadArea:=Range("A33:A38")SolverChangecellRef:=Range("F4:F6"), _ 
 relation:=1, _ 
 formulaText:=200 
SolverSolve userFinish:=False
```