# Workbook ReadOnly Property

## Business Description
Returns True if the object has been opened as read-only. Read-only Boolean.

## Behavior
ReturnsTrueif the object has been opened as read-only. Read-onlyBoolean.

## Example Usage
```vba
If ActiveWorkbook.ReadOnlyThen 
 ActiveWorkbook.SaveAs fileName:="NEWFILE.XLS" 
End If
```