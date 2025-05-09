# SortField SortOnValue Property

## Business Description
Retun the value on which the sort is performed for the specified SortField object. Read-only.

## Behavior
Retun the value on which the sort is performed for the specifiedSortFieldobject. Read-only.

## Example Usage
```vba
ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear 
ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add(Range("B1:B25"), _ 
 xlSortOnFontColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(0, 0, 0) 
 
With ActiveWorkbook.Worksheets("Sheet1").Sort 
 .SetRange Range("A1:B25") 
 .Header = xlGuess 
 .MatchCase = False 
 .Orientation = xlTopToBottom 
 .SortMethod = xlPinYin 
 .Apply 
End With
```