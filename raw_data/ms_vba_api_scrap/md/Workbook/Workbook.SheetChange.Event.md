# Workbook SheetChange Event

## Business Description
Occurs when cells in any worksheet are changed by the user or by an external link.

## Behavior
Occurs when cells in any worksheet are changed by the user or by an external link.

## Example Usage
```vba
Private Sub Workbook_SheetChange(ByVal Sh As Object, _ 
 ByVal Source As Range) 
 ' runs when a sheet is changed 
End Sub
```