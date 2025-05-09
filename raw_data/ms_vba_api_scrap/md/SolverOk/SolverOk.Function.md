# SolverOk Function

## Business Description
Defines a basic Solver model. Equivalent to clicking Solver in the Data | Analysis group and then specifying options in the Solver Parameters dialog box.

## Behavior
Defines a basic Solver model. Equivalent to clickingSolverin theData|Analysisgroup and then specifying options in theSolver Parametersdialog box.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
SolverReset 
SolverOptions precision:=0.001SolverOKSetCell:=Range("TotalProfit"), _ 
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