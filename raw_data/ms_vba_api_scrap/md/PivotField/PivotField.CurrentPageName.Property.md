# PivotField CurrentPageName Property

## Business Description
Returns or sets the currently displayed page of the specified PivotTable report. The name of the page appears in the page field. Note that this property works only if the currently displayed page already exists. Read/write String.

## Behavior
Returns or sets the currently displayed page of the specified PivotTable report. The name of the page appears in the page field. Note that this property works only if the currently displayed page already exists. Read/writeString.

## Example Usage
```vba
ActiveSheet.PivotTables("PivotTable1") _ 
 .PivotFields("[Customers]").CurrentPageName= _ 
 "[Customers].[All Customers].[USA]"
```