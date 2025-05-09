# CubeField ShowInFieldList Property

## Business Description
When set to True (default), a CubeField object will be shown in the field list. Read/write Boolean.

## Behavior
When set toTrue(default), aCubeFieldobject will be shown in the field list. Read/writeBoolean.

## Example Usage
```vba
Sub IsCubeFieldInList() 
 
 Dim pvtTable As PivotTable 
 Dim cbeField As CubeField 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set cbeField = pvtTable.CubeFields("[Country]") 
 
 ' Determine if a CubeField can be seen. 
 If cbeField.ShowInFieldList= True Then 
 MsgBox "The CubeField object can be seen in the field list." 
 Else 
 MsgBox "The CubeField object cannot be seen in the field list." 
 End If 
 
End Sub
```