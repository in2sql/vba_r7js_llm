# CustomView PrintSettings Property

## Business Description
True if print settings are included in the custom view. Read-only Boolean.

## Behavior
Trueif print settings are included in the custom view. Read-onlyBoolean.

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
 .Cells(rw, 2).Value = v.PrintSettings.Cells(rw, 3).Value = v.RowColSettings 
 Next 
End With
```