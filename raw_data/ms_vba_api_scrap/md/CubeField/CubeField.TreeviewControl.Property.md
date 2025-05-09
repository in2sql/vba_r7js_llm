# CubeField TreeviewControl Property

## Business Description
Returns the TreeviewControl object of the CubeField object, representing the cube manipulation control of an OLAP-based PivotTable report. Read-only.

## Behavior
Returns theTreeviewControlobject of theCubeFieldobject, representing the cube manipulation control of an OLAP-based PivotTable report. Read-only.

## Example Usage
```vba
ActiveSheet.PivotTables("PivotTable2") _ 
 .CubeFields(1).TreeviewControl.Drilled = _ 
 Array(Array("", ""), _ 
 Array("[state].[states].[CA]", _ 
 "[state].[states].[MD]"))
```