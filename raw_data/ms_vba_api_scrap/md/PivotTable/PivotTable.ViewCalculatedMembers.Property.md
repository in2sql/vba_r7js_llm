# PivotTable ViewCalculatedMembers Property

## Business Description
When set to True (default), calculated members for Online Analytical Processing (OLAP) PivotTables can be viewed. Read/write Boolean.

## Behavior
When set toTrue(default), calculated members for Online Analytical Processing (OLAP) PivotTables can be viewed. Read/writeBoolean.

## Example Usage
```vba
Sub CheckViewCalculatedMembers() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Determine if calculated members can be viewed. 
 If pvtTable.ViewCalculatedMembers= True Then 
 MsgBox "Calculated members can be viewed." 
 Else 
 MsgBox "Calculated members cannot be viewed." 
 End If 
 
End Sub
```