# PivotFilters Add2 Method

## Business Description
Adds new filters to the PivotFilters collection.

## Behavior
Adds new filters to thePivotFilterscollection.

## Example Usage
```vba
ActiveCell.PivotField.PivotFilters.Add FilterType := xlThisWeek 
 
ActiveCell.PivotField.PivotFilters.Add FilterType := xlTopCount DataField := MyPivotField2 Value1 := 10 
 
ActiveCell.PivotField.PivotFilters.Add FilterType := xlCaptionIsNotBetween Value1 := "A" Value2 := "G" 
 
ActiveCell.PivotField.PivotFilters.Add FilterType := xlValueIsGreaterThanOrEqualTo DataField := MyPivotField2 Value1 := 10000
```