# SolverSave Function

## Business Description
Saves the Solver problem specifications on the worksheet.

## Behavior
Saves the Solver problem specifications on the worksheet.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
SolverReset 
SolverOptions Precision:=0.001 
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
SolverSolve UserFinish:=FalseSolverSaveSaveArea:=Range("A33")
```