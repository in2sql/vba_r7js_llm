# OLEDBConnection Connection Property

## Business Description
Returns or sets a string that contains OLE DB settings that enable Microsoft Excel to connect to an OLE DB data source. Read/write Variant.

## Behavior
Returns or sets a string that contains OLE DB settings that enable Microsoft Excel to connect to an OLE DB data source. Read/writeVariant.

## Example Usage
```vba
With ActiveWorkbook.PivotCaches.Add(SourceType:=xlExternal) 
 .Connection= _ 
 "OLEDB;Provider=MSOLAP;Location=srvdata;Initial Catalog=National" 
 .MaintainConnection = True 
 .CreatePivotTable TableDestination:=Range("A3"), _ 
 TableName:= "PivotTable1" 
End With 
With ActiveSheet.PivotTables("PivotTable1") 
 .SmallGrid = False 
 .PivotCache.RefreshPeriod = 0 
 With .CubeFields("[state]") 
 .Orientation = xlColumnField 
 .Position = 0 
 End With 
 With .CubeFields("[Measures].[Count Of au_id]") 
 .Orientation = xlDataField 
 .Position = 0 
 End With 
End With
```