# SolverAdd Function

## Business Description
Adds a constraint to the current problem. Equivalent to clicking Solver in the Data | Analysis group and then clicking Add in the Solver Parameters dialog box.

## Behavior
Adds a constraint to the current problem. Equivalent to clickingSolverin theData|Analysisgroup and then clickingAddin theSolver Parametersdialog box.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
SolverReset 
SolverOptions precision:=0.001 
SolverOK setCell:=Range("TotalProfit"), _ 
 maxMinVal:=1, _ 
 byChange:=Range("C4:E6")SolverAddcellRef:=Range("F4:F6"), _ 
 relation:=1, _ 
 formulaText:=100SolverAddcellRef:=Range("C4:E6"), _ 
 relation:=3, _ 
 formulaText:=0SolverAddcellRef:=Range("C4:E6"), _ 
 relation:=4 
SolverSolve userFinish:=False 
SolverSave saveArea:=Range("A33")
```