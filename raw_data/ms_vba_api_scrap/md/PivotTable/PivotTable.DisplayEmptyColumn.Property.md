# PivotTable DisplayEmptyColumn Property

## Business Description
Returns True when the non-empty MDX keyword is included in the query to the OLAP provider for the value axis. The OLAP provider will not return empty columns in the result set. Returns False when the non-empty keyword is omitted. Read/write Boolean.

## Behavior
ReturnsTruewhen the non-empty MDX keyword is included in the query to the OLAP provider for the value axis. The OLAP provider will not return empty columns in the result set. ReturnsFalsewhen the non-empty keyword is omitted. Read/writeBoolean.

## Example Usage
```vba
Sub CheckSetting() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Determine display setting for empty columns. 
 If pvtTable.DisplayEmptyColumn= False Then 
 MsgBox "Empty columns will not be displayed." 
 Else 
 MsgBox "Empty columns will be displayed." 
 End If 
 
End Sub
```