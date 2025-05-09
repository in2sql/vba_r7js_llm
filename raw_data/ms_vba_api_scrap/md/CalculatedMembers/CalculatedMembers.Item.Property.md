# CalculatedMembers Item Property

## Business Description
Returns a single object from a collection.

## Behavior
Returns a single object from a collection.

## Example Usage
```vba
Sub CheckValidity() 
 
 Dim pvtTable As PivotTable 
 Dim pvtCache As PivotCache 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set pvtCache = Application.ActiveWorkbook.PivotCaches.Item(1) 
 
 ' Handle run-time error if external source is not an OLEDB data source. 
 On Error GoTo Not_OLEDB 
 
 ' Check connection setting and make connection if necessary. 
 If pvtCache.IsConnected = False Then 
 pvtCache.MakeConnection 
 End If 
 
 ' Check if calculated member is valid. 
 If pvtTable.CalculatedMembers.Item(1).IsValid = True Then 
 MsgBox "The calculated member is valid." 
 Else 
 MsgBox "The calculated member is not valid." 
 End If 
 
Not_OLEDB: 
MsgBox "The source is not an OLEDB data source." 
Exit Sub 
 
End Sub
```