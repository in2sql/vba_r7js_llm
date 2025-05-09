# PivotCache CreatePivotChart Method

## Business Description
Creates a standalone PivotChart from a PivotCache object. A Shape object is returned.

## Behavior
Creates a standalone PivotChart from aPivotCache Object (Excel)object. AShape Object (Excel)object is returned.

## Example Usage
```vba
Workbooks("Book1").Connections.Add _
     "cubes4 Adventure Works DW 2008 Special Char Adventure Works", "", Array( _
     "OLEDB;Provider=MSOLAP.4;Integrated Security=SSPI;Persist Security Info=True;Data Source=<server name here>;Initial Catalog=Adventure Works DW 2008" _
     , " Special Char"), Array("Adventure Works"), 1
   ActiveWorkbook.PivotCaches.Create(SourceType:=xlExternal, SourceData:= _
     ActiveWorkbook.Connections( _
     "cubes4 Adventure Works DW 2008 Special Char Adventure Works"), Version:= _
     xlPivotTableVersion14).CreatePivotChart(ChartDestination:="Sheet1").Select

   ActiveChart.ChartType = xlColumnClustered
```