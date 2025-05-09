# SolverDelete Function

## Business Description
Deletes an existing constraint. Equivalent to clicking Solver in the Data | Analysis group and then clicking Delete in the Solver Parameters dialog box.

## Behavior
Deletes an existing constraint. Equivalent to clickingSolverin theData|Analysisgroup and then clickingDeletein theSolver Parametersdialog box.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
SolverLoad loadArea:=Range("A33:A38")SolverDeletecellRef:=Range("C4:E6"), _ 
 relation:=4 
SolverSolve userFinish:=False
```