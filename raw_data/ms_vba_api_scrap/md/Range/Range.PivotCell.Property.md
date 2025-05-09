# Range PivotCell Property

## Business Description
Returns a PivotCell object that represents a cell in a PivotTable report.

## Behavior
Returns aPivotCellobject that represents a cell in a PivotTable report.

## Example Usage
```vba
Sub CheckPivotCell() 
 
 'Determine the name of the PivotTable the PivotCell is located in. 
 MsgBox "Cell A3 is located in PivotTable: " & _ 
 Application.Range("A3").PivotCell.Parent 
 
End Sub
```