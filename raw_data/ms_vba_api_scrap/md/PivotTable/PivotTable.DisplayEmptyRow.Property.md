# PivotTable DisplayEmptyRow Property

## Business Description
Returns True when the non-empty MDX keyword is included in the query to the OLAP provider for the category axis. The OLAP provider will not return empty rows in the result set. Returns False when the non-empty keyword is omitted. Read/write Boolean.

## Behavior
ReturnsTruewhen the non-empty MDX keyword is included in the query to the OLAP provider for the category axis. The OLAP provider will not return empty rows in the result set. ReturnsFalsewhen the non-empty keyword is omitted. Read/writeBoolean.

## Example Usage
```vba
Sub CheckSetting() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Determine display setting for empty rows. 
 If pvtTable.DisplayEmptyRow= False Then 
 MsgBox "Empty rows will not be displayed." 
 Else 
 MsgBox "Empty rows will be displayed." 
 End If 
 
End Sub
```