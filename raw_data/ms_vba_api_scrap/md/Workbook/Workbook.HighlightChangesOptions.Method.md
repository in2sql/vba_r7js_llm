# Workbook HighlightChangesOptions Method

## Business Description
Controls how changes are shown in a shared workbook.

## Behavior
Controls how changes are shown in a shared workbook.

## Example Usage
```vba
With ActiveWorkbook 
 .HighlightChangesOptions_ 
 When:=xlSinceMyLastSave, _ 
 Who:="Everyone" 
 .ListChangesOnNewSheet = True 
End With
```