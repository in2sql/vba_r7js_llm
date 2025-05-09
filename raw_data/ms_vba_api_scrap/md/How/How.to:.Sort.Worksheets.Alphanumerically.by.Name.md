# How to: Sort Worksheets Alphanumerically by Name

## Business Description
The following example shows how to sort the worksheets in a workbook alphanumerically based on the name of the sheet by using the Name property of the Worksheet object.

## Behavior
The following example shows how to sort the worksheets in a workbook alphanumerically based on the name of the sheet by using theNameproperty of theWorksheetobject.

## Example Usage
```vba
Sub SortSheetsTabName()
    Application.ScreenUpdating = False
    Dim iSheets%, i%, j%
    iSheets = Sheets.Count
    For i = 1 To iSheets - 1
        For j = i + 1 To iSheets
            If Sheets(j).Name < Sheets(i).Name Then
                Sheets(j).Move before:=Sheets(i)
            End If
        Next j
    Next i
    Application.ScreenUpdating = True
End Sub
```