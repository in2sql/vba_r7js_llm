# SolverOkDialog Function

## Business Description
Same as the SolverOK function, but also displays the Solver dialog box.

## Behavior
Same as theSolverOKfunction, but also displays theSolverdialog box.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
SolverLoad LoadArea:=Range("A33:A38") 
SolverResetSolverOKDialogSetCell:=Range("TotalProfit") 
SolverSolve UserFinish:=False
```