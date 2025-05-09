# Workbook ListChangesOnNewSheet Property

## Business Description
True if changes to the shared workbook are shown on a separate worksheet. Read/write Boolean.

## Behavior
Trueif changes to the shared workbook are shown on a separate worksheet. Read/writeBoolean.

## Example Usage
```vba
With ActiveWorkbook 
 .HighlightChangesOptions _ 
 When:=xlSinceMyLastSave, _ 
 Who:="Everyone" 
 .ListChangesOnNewSheet= True 
End With
```