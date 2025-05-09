# PivotTable VisualTotals Property

## Business Description
True (default) to enable Online Analytical Processing (OLAP) PivotTables to retotal after an item has been hidden from view. Read/write Boolean.

## Behavior
True(default) to enable Online Analytical Processing (OLAP) PivotTables to retotal after an item has been hidden from view. Read/writeBoolean.

## Example Usage
```vba
Sub CheckVisualTotals() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Determine if visual totals is enabled for OLAP PivotTables. 
 If pvtTable.VisualTotals= True Then 
 MsgBox "Ability enabled to re-total after an item " & _ 
 "has been hidden from view." 
 Else 
 MsgBox "Unable to re-total items not hidden from view." 
 End If 
 
End Sub
```