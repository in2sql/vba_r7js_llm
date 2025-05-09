# PivotCache Recordset Property

## Business Description
Returns or sets a Recordset object that's used as the data source for the specified PivotTable cache. Read/write.

## Behavior
Returns or sets aRecordsetobject that's used as the data source for the specified PivotTable cache. Read/write.

## Example Usage
```vba
Dim cnnConn As ADODB.Connection 
Dim rstRecordset As ADODB.Recordset 
Dim cmdCommand As ADODB.Command 
 
' Open the connection. 
Set cnnConn = New ADODB.Connection 
With cnnConn 
 .ConnectionString = _ 
 "Provider=Microsoft.Jet.OLEDB.4.0" 
 .Open "C:\perfdate\record.mdb" 
End With 
 
' Set the command text. 
Set cmdCommand = New ADODB.Command 
Set cmdCommand.ActiveConnection = cnnConn 
With cmdCommand 
 .CommandText = "Select Speed, Pressure, Time From DynoRun" 
 .CommandType = adCmdText 
 .Execute 
End With 
 
' Open the recordset. 
Set rstRecordset = New ADODB.Recordset 
Set rstRecordset.ActiveConnection = cnnConn 
rstRecordset.Open cmdCommand 
 
' Create a PivotTable cache and report. 
Set objPivotCache = ActiveWorkbook.PivotCaches.Add( _ 
 SourceType:=xlExternal) 
Set objPivotCache.Recordset= rstRecordset 
With objPivotCache 
 .CreatePivotTable TableDestination:=Range("A3"), _ 
 TableName:="Performance" 
End With 
 
With ActiveSheet.PivotTables("Performance") 
 .SmallGrid = False 
 With .PivotFields("Pressure") 
 .Orientation = xlRowField 
 .Position = 1 
 End With 
 With .PivotFields("Speed") 
 .Orientation = xlColumnField 
 .Position = 1 
 End With 
 With .PivotFields("Time") 
 .Orientation = xlDataField 
 .Position = 1 
 End With 
End With 
 
' Close the connections and clean up. 
cnnConn.Close 
Set cmdCommand = Nothing 
Set rstRecordSet = Nothing 
Set cnnConn = Nothing
```