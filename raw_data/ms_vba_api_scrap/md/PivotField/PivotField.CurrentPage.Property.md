# PivotField CurrentPage Property

## Business Description
Returns or sets the current page showing for the page field (valid only for page fields). Read/write PivotItem.

## Behavior
Returns or sets the current page showing for the page field (valid only for page fields). Read/writePivotItem.

## Example Usage
```vba
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
strPgName = pvtTable.PivotFields("Country").CurrentPage.Name
```