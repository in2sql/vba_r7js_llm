# SolverOptions Function

## Business Description
Allows you to specify advanced options for your Solver model. This function and its arguments correspond to the options in the Solver Options dialog box.

## Behavior
Allows you to specify advanced options for your Solver model. This function and its arguments correspond to the options in theSolver Optionsdialog box.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
SolverResetSolverOptionsPrecision:=0.001 
SolverOK SetCell:=Range("TotalProfit"), _ 
    MaxMinVal:=1, _ 
    ByChange:=Range("C4:E6") 
SolverAdd CellRef:=Range("F4:F6"), _ 
    Relation:=1, _ 
    FormulaText:=100 
SolverAdd CellRef:=Range("C4:E6"), _ 
    Relation:=3, _ 
    FormulaText:=0 
SolverAdd CellRef:=Range("C4:E6"), _ 
    Relation:=4 
SolverSolve UserFinish:=False 
SolverSave SaveArea:=Range("A33")
```