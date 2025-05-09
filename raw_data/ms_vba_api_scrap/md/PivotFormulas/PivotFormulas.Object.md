# PivotFormulas Object

## Business Description
Represents the collection of formulas for a PivotTable report. Each formula is represented by a PivotFormula object.

## Behavior
Represents the collection of formulas for a PivotTable report. Each formula is represented by aPivotFormulaobject.

## Example Usage
```vba
For Each pf in ActiveSheet.PivotTables(1).PivotFormulasCells(r, 1).Value = pf.Formula 
 r = r + 1 
Next
```