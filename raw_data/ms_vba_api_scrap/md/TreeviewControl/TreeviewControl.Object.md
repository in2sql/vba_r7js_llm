# TreeviewControl Object

## Business Description
Represents the hierarchical member-selection control of a cube field.

## Behavior
Represents the hierarchical member-selection control of a cube field.

## Example Usage
```vba
ActiveSheet.PivotTables("PivotTable2") _ 
 .CubeFields(1).TreeviewControl.Drilled = _ 
 Array(Array("", ""), _ 
 Array("[state].[states].[CA]", _ 
 "[state].[states].[MD]"))
```