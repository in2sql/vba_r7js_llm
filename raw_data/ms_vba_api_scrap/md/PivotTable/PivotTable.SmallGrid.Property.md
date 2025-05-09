# PivotTable SmallGrid Property

## Business Description
True if Microsoft Excel uses a grid that's two cells wide and two cells deep for a newly created PivotTable report. False if Excel uses a blank stencil outline. Read/write Boolean.

## Behavior
Trueif Microsoft Excel uses a grid that's two cells wide and two cells deep for a newly created PivotTable report.Falseif Excel uses a blank stencil outline. Read/writeBoolean.

## Example Usage
```vba
With ActiveWorkbook.PivotCaches.Add(SourceType:=xlExternal) 
 .Connection = _ 
 "OLEDB;Provider=MSOLAP;Location=srvdata;Initial Catalog=National" 
 .MaintainConnection = True 
 .CreatePivotTable TableDestination:=Range("A3"), _ 
 TableName:= "PivotTable1" 
End With 
With ActiveSheet.PivotTables("PivotTable1") 
 .SmallGrid= False 
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