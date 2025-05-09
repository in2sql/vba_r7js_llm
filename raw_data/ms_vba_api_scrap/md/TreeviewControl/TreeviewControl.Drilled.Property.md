# TreeviewControl Drilled Property

## Business Description
Sets the "drilled" (expanded, or visible) status of the cube field members in the hierarchical member-selection control of a cube field. This property is used primarily for macro recording and isn't intended for any other use. Read/write.

## Behavior
Sets the "drilled" (expanded, or visible) status of the cube field members in the hierarchical member-selection control of a cube field. This property is used primarily for macro recording and isn't intended for any other use. Read/write.

## Example Usage
```vba
ActiveSheet.PivotTables("PivotTable1").CubeFields(1) _ 
 .TreeviewControl.Drilled= _ 
 Array(Array("", "", "", "", "", "", "", "", _ 
 "", "", "", ""), _ 
 Array("[state].[states].[AB]", _ 
 "[state].[states].[CA]", _ 
 "[state].[states].[IN]", _ 
 "[state].[states].[KS]", _ 
 "[state].[states].[KY]", _ 
 "[state].[states].[MD]", _ 
 "[state].[states].[MI]", _ 
 "[state].[states].[OH]", _ 
 "[state].[states].[OR]", _ 
 "[state].[states].[TN]", _ 
 "[state].[states].[UT]", _ 
 "[state].[states].[WA]"))
```