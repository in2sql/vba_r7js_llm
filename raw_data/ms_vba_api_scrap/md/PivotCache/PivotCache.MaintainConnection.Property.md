# PivotCache MaintainConnection Property

## Business Description
True if the connection to the specified data source is maintained after the refresh and until the workbook is closed. The default value is True. Read/write Boolean.

## Behavior
Trueif the connection to the specified data source is maintained after the refresh and until the workbook is closed. The default value isTrue. Read/writeBoolean.

## Example Usage
```vba
With ActiveWorkbook.PivotCaches.Add(SourceType:=xlExternal) 
 .Connection = _ 
 "OLEDB;Provider=MSOLAP;Location=srvdata;Initial Catalog=National" 
 .MaintainConnection= False 
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