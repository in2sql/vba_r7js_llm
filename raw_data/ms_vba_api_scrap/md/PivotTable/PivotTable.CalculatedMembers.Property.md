# PivotTable CalculatedMembers Property

## Business Description
Returns a CalculatedMembers collection representing all the calculated members and calculated measures for an OLAP PivotTable.

## Behavior
Returns aCalculatedMemberscollection representing all the calculated members and calculated measures for an OLAP PivotTable.

## Example Usage
```vba
Sub UseCalculatedMember() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Add the calculated member. 
 pvtTable.CalculatedMembers.Add Name:="[Beef]", _ 
 Formula:="'{[Product].[All Products].Children}'", _ 
 Type:=xlCalculatedSet 
 
End Sub
```