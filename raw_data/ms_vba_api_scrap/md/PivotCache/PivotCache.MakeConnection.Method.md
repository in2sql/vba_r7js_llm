# PivotCache MakeConnection Method

## Business Description
Establishes a connection for the specified PivotTable cache.

## Behavior
Establishes a connection for the specified PivotTable cache.

## Example Usage
```vba
Sub UseMakeConnection() 
 
    Dim pvtCache As PivotCache 
 
    Set pvtCache = Application.ActiveWorkbook.PivotCaches.Item(1) 
 
    ' Handle run-time error if external source is not an OLE DB data source. 
    On Error GoTo Not_OLEDB 
 
    ' Check connection setting and make connection if necessary. 
    If pvtCache.IsConnected = True Then 
        MsgBox "The MakeConnection method is not needed." 
    Else 
        pvtCache.MakeConnection 
        MsgBox "A connection has been made." 
    End If 
    Exit Sub 
 
Not_OLEDB: 
    MsgBox "The data source is not an OLE DB data source" 
 
End Sub
```