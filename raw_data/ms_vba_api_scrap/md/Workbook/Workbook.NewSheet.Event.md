# Workbook NewSheet Event

## Business Description
Occurs when a new sheet is created in the workbook.

## Behavior
Occurs when a new sheet is created in the workbook.

## Example Usage
```vba
Private Sub Workbook_NewSheet(ByVal Sh as Object) 
 Sh.Move After:= Sheets(Sheets.Count) 
End Sub
```