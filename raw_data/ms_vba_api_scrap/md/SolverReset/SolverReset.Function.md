# SolverReset Function

## Business Description
Resets all cell selections and constraints in the Solver Parameters dialog box and restores all the settings in the Solver Options dialog box to their defaults. Equivalent to clicking Reset All in the Solver Parameters dialog box.

## Behavior
Resets all cell selections and constraints in theSolver Parametersdialog box and restores all the settings in theSolver Optionsdialog box to their defaults. Equivalent to clickingReset Allin theSolver Parametersdialog box. TheSolverResetfunction is called automatically when you call theSolverLoadfunction, if theMergeargument isFalseor omitted.

## Example Usage
```vba
Worksheets("Sheet1").ActivateSolverResetSolverOptions Precision:=0.001 
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