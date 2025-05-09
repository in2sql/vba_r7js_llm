# Workbook Saved Property

## Business Description
True if no changes have been made to the specified workbook since it was last saved. Read/write Boolean.

## Behavior
Trueif no changes have been made to the specified workbook since it was last saved. Read/writeBoolean.

## Example Usage
```vba
If Not ActiveWorkbook.SavedThen 
 MsgBox "This workbook contains unsaved changes." 
End If
```