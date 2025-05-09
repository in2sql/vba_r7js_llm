# Range Row Property

## Business Description
Returns the number of the first row of the first area in the range. Read-only Long.

## Behavior
Returns the number of the first row of the first area in the range. Read-onlyLong.

## Example Usage
```vba
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    'If the double click occurs on the header row or an empty cell, exit the macro.
    If Target.Row = 1 Then Exit Sub
    If Target.Row > ActiveSheet.UsedRange.Rows.Count Then Exit Sub
    If Target.Column > ActiveSheet.UsedRange.Columns.Count Then Exit Sub
    
    'Override the default double-click behavior with this function.
    Cancel = True
    
    'Declare your variables.
    Dim wks As Worksheet, xRow As Long
    
    'If an error occurs, use inline error handling.
    On Error Resume Next
    
    'Set the target worksheet as the worksheet whose name is listed in the first cell of the current row.
    Set wks = Worksheets(CStr(Cells(Target.Row, 1).Value))
    'If there is an error, exit the macro.
    If Err > 0 Then
        Err.Clear
        Exit Sub
    'Otherwise, find the next empty row in the target worksheet and copy the data into that row.
    Else
        xRow = wks.Cells(wks.Rows.Count, 1).End(xlUp).Row + 1
        wks.Range(wks.Cells(xRow, 1), wks.Cells(xRow, 7)).Value = _
        Range(Cells(Target.Row, 1), Cells(Target.Row, 7)).Value
    End If
End Sub
```