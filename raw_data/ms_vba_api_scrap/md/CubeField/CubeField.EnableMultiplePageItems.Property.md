# CubeField EnableMultiplePageItems Property

## Business Description
Set to True to allow multiple items in the page field area for OLAP PivotTables to be selected. The default value is False. Read/write Boolean.

## Behavior
Set toTrueto allow multiple items in the page field area for OLAP PivotTables to be selected. The default value isFalse. Read/writeBoolean.

## Example Usage
```vba
Sub UseMultiplePageItems() 
 
 Dim pvtTable As PivotTable 
 Dim cbeField As CubeField 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set cbeField = pvtTable.CubeFields("[Country]") 
 
 ' Determine setting for mulitple page items. 
 If cbeField.EnableMultiplePageItems= False Then 
 MsgBox "Mulitple page items cannot be selected." 
 Else 
 MsgBox "Multiple page items can be selected." 
 End If 
End Sub
```