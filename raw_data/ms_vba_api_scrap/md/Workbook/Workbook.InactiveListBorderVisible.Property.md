# Workbook InactiveListBorderVisible Property

## Business Description
A Boolean value that specifies whether list borders are visible when a list is not active. Returns True if the border is visible. Read/write Boolean.

## Behavior
ABooleanvalue that specifies whether list borders are visible when a list is not active. ReturnsTrueif the border is visible. Read/writeBoolean.

## Example Usage
```vba
Sub HideListBorders() 
 
 ActiveWorkbook.InactiveListBorderVisible= False 
 
End Sub
```