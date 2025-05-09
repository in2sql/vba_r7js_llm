# Slicers Add Method

## Business Description
Creates a new slicer and returns a Slicer object.

## Behavior
Creates a new slicer and returns aSlicerobject.

## Example Usage
```vba
Sub CreateNewSlicer() 
 ActiveWorkbook.SlicerCaches.Add("Adventure Works", _ 
 "[Customer].[Customer Geography]").Slicers.Add ActiveSheet, _ 
 "[Customer].[Customer Geography].[Country]", "Country 1", "Country", _ 
 252, 522, 144, 216) 
End Sub
```