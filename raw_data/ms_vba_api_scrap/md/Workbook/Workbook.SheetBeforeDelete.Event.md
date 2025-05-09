# Workbook SheetBeforeDelete Event

## Business Description
Occurs when any sheet is deleted.

## Behavior
Occurs when any sheet is deleted.

## Example Usage
```vba
Private Sub Workbook_SheetBeforeDelete(ByVal Sh As Object) 
 MsgBox Sh.Name 
End Sub
```