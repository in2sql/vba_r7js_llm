# PivotValueCell Value Property

## Business Description
Returns the value at the location. The value is the value after ShowAs and other calculations have been applied. Variant can be Empty, Number, Date, String, or Error value.

## Behavior
Returns the value at the location. The value is the value afterShowAsand other calculations have been applied. Variant can beEmpty,Number,Date,String, orErrorvalue.

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