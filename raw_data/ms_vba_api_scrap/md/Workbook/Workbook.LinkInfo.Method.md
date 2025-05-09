# Workbook LinkInfo Method

## Business Description
Returns the link date and update status.

## Behavior
Returns the link date and update status.

## Example Usage
```vba
If ActiveWorkbook.LinkInfo( _ 
 "Word.Document|Document1!'!DDE_LINK1", xlUpdateState, _ 
 xlOLELinks) = 1 Then 
 MsgBox "Link updates automatically" 
End If
```