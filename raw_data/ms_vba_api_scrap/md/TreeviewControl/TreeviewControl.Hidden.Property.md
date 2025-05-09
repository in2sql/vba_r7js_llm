# TreeviewControl Hidden Property

## Business Description
Returns or sets a Variant value that represents the hidden status of the cube field members in the hierarchical member selection control of a cube field.

## Behavior
Returns or sets a Variant value that represents the hidden status of the cube field members in the hierarchical member selection control of a cube field.

## Example Usage
```vba
ActiveSheet.PivotTables("PivotTable1").CubeFields(1) _ 
 .TreeviewControl.Hidden = _ 
 Array(Array(""), Array(""), _ 
 Array("[state].[states].[CA].[Covelo]"))
```