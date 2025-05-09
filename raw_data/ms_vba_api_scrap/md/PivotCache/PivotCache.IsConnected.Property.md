# PivotCache IsConnected Property

## Business Description
Returns True if the MaintainConnection property is True and the PivotTable cache is currently connected to its source. Returns False if it is not currently connected to its source. Read-only Boolean.

## Behavior
ReturnsTrueif theMaintainConnectionproperty isTrueand the PivotTable cache is currently connected to its source. ReturnsFalseif it is not currently connected to its source. Read-onlyBoolean.

## Example Usage
```vba
Sub CheckIsConnected() 
 
 ' Handle run-time error if external source is not OLE DB. 
 On Error GoTo Not_OLEDB 
 
 ' Check connection setting and notify the user accordingly. 
 If Application.ActiveWorkbook.PivotCaches.Item(1).IsConnected = True Then 
 MsgBox "The PivotCache is currently connected to its source." 
 Else 
 MsgBox "The PivotCache is not currently connected to its source." 
 End If 
 Exit Sub 
 
Not_OLEDB: 
 MsgBox "The data source is not an OLE DB data source." 
 
End Sub
```