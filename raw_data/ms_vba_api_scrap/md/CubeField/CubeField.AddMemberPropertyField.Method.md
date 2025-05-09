# CubeField AddMemberPropertyField Method

## Business Description
Adds a member property field to the display for the cube field.

## Behavior
Adds a member property field to the display for the cube field.

## Example Usage
```vba
Sub UseAddMemberPropertyField() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 With pvtTable 
 .ManualUpdate = True 
 .CubeFields("[Country]").LayoutForm = xlOutline 
 .CubeFields("[Country]").AddMemberPropertyField_ 
 Property:="[Country].[Area].[Description]" 
 .ManualUpdate = False 
 End With 
 
End Sub
```