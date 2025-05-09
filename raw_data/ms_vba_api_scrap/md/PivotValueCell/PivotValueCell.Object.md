# PivotValueCell Object

## Business Description
Provides a way to expose values of cells in the case that actual cells (Range objects) are not available.

## Behavior
Provides a way to expose values of cells in the case that actual cells (Rangeobjects) are not available.

## Example Usage
```vba
Sub TestEquality()
Dim X As Double
Dim Y As Double

'This code assumes that you have a Standalone PivotChart on one of the worksheets.
X = ThisWorkbook.PivotTables(1).PivotValueCell(1, 1).Value
Y = ThisWorkbook.PivotTables(1).PivotValueCell(1, 2).Value

If X > Y Then
MsgBox "X is greater than Y"
Else
MsgBox "Y is greater than X"
End If
End Sub
```