# CustomView RowColSettings Property

## Business Description
True if the custom view includes settings for hidden rows and columns (including filter information). Read-only Boolean.

## Behavior
Trueif the custom view includes settings for hidden rows and columns (including filter information). Read-onlyBoolean.

## Example Usage
```vba
With Worksheets(1) 
 .Cells(1,1).Value = "Name" 
 .Cells(1,2).Value = "Print Settings" 
 .Cells(1,3).Value = "RowColSettings" 
 rw = 0 
 For Each v In ActiveWorkbook.CustomViews 
 rw = rw + 1 
 .Cells(rw, 1).Value = v.Name 
 .Cells(rw, 2).Value = v.PrintSettings 
 .Cells(rw, 3).Value = v.RowColSettingsNext 
End With
```