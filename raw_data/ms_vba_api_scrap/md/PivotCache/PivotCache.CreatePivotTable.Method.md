# PivotCache CreatePivotTable Method

## Business Description
Creates a PivotTable report based on a PivotCache object. Returns a PivotTable object.

## Behavior
Creates a PivotTable report based on aPivotCacheobject. Returns aPivotTableobject.

## Example Usage
```vba
With ActiveWorkbook.PivotCaches.Add(SourceType:=xlExternal) 
 .Connection = _ 
 "OLEDB;Provider=MSOLAP;Location=srvdata;Initial Catalog=National" 
 .CommandType = xlCmdCube 
 .CommandText = Array("Sales") 
 .MaintainConnection = True 
 .CreatePivotTableTableDestination:=Range("A3"), _ 
 TableName:= "PivotTable1" 
End With 
With ActiveSheet.PivotTables("PivotTable1") 
 .SmallGrid = False 
 .PivotCache.RefreshPeriod = 0 
 With .CubeFields("[state]") 
 .Orientation = xlColumnField 
 .Position = 1 
 End With 
 With .CubeFields("[Measures].[Count Of au_id]") 
 .Orientation = xlDataField 
 .Position = 1 
 End With 
End With
```