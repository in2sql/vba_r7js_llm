# Names Object

## Business Description
A collection of all the Name objects in the application or workbook.

## Behavior
A collection of all theNameobjects in the application or workbook.

## Example Usage
```vba
Set nms = ActiveWorkbook.Names 
Set wks = Worksheets(1) 
For r = 1 To nms.Count 
    wks.Cells(r, 2).Value = nms(r).Name 
    wks.Cells(r, 3).Value = nms(r).RefersToRange.Address 
Next
```