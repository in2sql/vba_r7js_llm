# CalculatedMember IsValid Property

## Business Description
Returns a Boolean that indicates whether the specified calculated member has been successfully instantiated with the OLAP provider during the current session.

## Behavior
Returns a Boolean that indicates whether the specified calculated member has been successfully instantiated with the OLAP provider during the current session.

## Example Usage
```vba
Sub CheckValidity() 
 
 Dim pvtTable As PivotTable 
 Dim pvtCache As PivotCache 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set pvtCache = Application.ActiveWorkbook.PivotCaches.Item(1) 
 
 ' Make connection for PivotTable before testing IsValid property. 
 pvtCache.MakeConnection 
 
 ' Check if calculated member is valid. 
 If pvtTable.CalculatedMembers.Item(1).IsValid= True Then 
 MsgBox "The calculated member is valid." 
 Else 
 MsgBox "The calculated member is not valid." 
 End If 
 
End Sub
```