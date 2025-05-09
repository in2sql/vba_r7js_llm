# SolverLoad Function

## Business Description
Loads existing Solver model parameters that have been saved to the worksheet.

## Behavior
Loads existing Solver model parameters that have been saved to the worksheet.

## Example Usage
```vba
Worksheets("Sheet1").ActivateSolverLoadloadArea:=Range("A33:A38") 
SolverChange cellRef:=Range("F4:F6"), _ 
 relation:=1, _ 
 formulaText:=200 
SolverSolve userFinish:=False
```