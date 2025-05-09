# PivotTable PivotValueCell Method

## Business Description
Retrieve the PivotValueCell object for a given PivotTable provided certain row and column indices.

## Behavior
Retrieve thePivotValueCell Object (Excel)object for a given PivotTable provided certain row and column indices.

## Example Usage
```vba
Sub TestEquality()
Dim X As Double
Dim Y As Double

'This code assumes you have a Standalone PivotChart on one of the worksheets
X = ThisWorkbook.PivotTables(1).PivotValueCell(1, 1).Value
Y = ThisWorkbook.PivotTables(1).PivotValueCell(1, 2).Value

If X > Y Then
MsgBox "X is greater than Y"
Else
MsgBox "Y is greater than X"
End If
End Sub
```