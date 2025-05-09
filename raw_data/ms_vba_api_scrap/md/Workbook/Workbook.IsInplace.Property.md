# Workbook IsInplace Property

## Business Description
True if the specified workbook is being edited in place. False if the workbook has been opened in Microsoft Excel for editing. Read-only Boolean.

## Behavior
Trueif the specified workbook is being edited in place.Falseif the workbook has been opened in Microsoft Excel for editing. Read-onlyBoolean.

## Example Usage
```vba
Private Sub Workbook_Open() 
 If ThisWorkbook.IsInPlace= True Then 
 MsgBox "Editing in place" 
 Else 
 MsgBox "Editing in Microsoft Excel" 
 End If 
End Sub
```