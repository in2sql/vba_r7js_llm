# PivotTable PivotFormulas Property

## Business Description
Returns a PivotFormulas object that represents the collection of formulas for the specified PivotTable report. Read-only.

## Behavior
Returns aPivotFormulasobject that represents the collection of formulas for the specified PivotTable report. Read-only.

## Example Usage
```vba
For Each pf in ActiveSheet.PivotTables(1).PivotFormulasr = r + 1 
 Cells(r, 1).Value = pf.Formula 
Next
```