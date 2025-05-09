# PivotCache Connection Property

## Business Description
Returns or sets a string that contains one of the following: OLE DB settings that enable Microsoft Excel to connect to an OLE DB data source; ODBC settings that enable Microsoft Excel to connect to an ODBC data source; a URL that enables Microsoft Excel to

## Behavior
Returns or sets a string that contains one of the following: OLE DB settings that enable Microsoft Excel to connect to an OLE DB data source; ODBC settings that enable Microsoft Excel to connect to an ODBC data source; a URL that enables Microsoft Excel to connect to a Web data source; the path to and file name of a text file, or the path to and file name of a file that specifies a database or Web query. Read/writeVariant.

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