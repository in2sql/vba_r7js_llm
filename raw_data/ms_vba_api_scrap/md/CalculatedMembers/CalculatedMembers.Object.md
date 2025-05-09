# CalculatedMembers Object

## Business Description
A collection of all the CalculatedMember objects on the specified PivotTable.

## Behavior
A collection of all theCalculatedMemberobjects on the specified PivotTable.

## Example Usage
```vba
Sub UseCalculatedMember() 
 Dim pvtTable As PivotTable 
 Set pvtTable = ActiveSheet.PivotTables(1)
 pvtTable.CalculatedMembers.Add Name:="[Beef]", _ 
 Formula:="'{[Product].[All Products].Children}'", _ 
 Type:=xlCalculatedSet 
 
End Sub
```