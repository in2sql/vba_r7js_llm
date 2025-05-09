# PivotCaches Create Method

## Business Description
Creates a new PivotCache.

## Behavior
Creates a new PivotCache.

## Example Usage
```vba
Workbooks("Book1").Connections.Add2 _
        "Target Connection Name", "", Array("OLEDB;Provider=MSOLAP.5;Integrated Security=SSPI;Persist Security Info=True;Data Source=##TargetServer##;Initial Catalog=Adventure Works DW", ""), "Adventure Works", 1
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlExternal, SourceData:=ActiveWorkbook.Connections("Target Connection Name"), _ Version:=xlPivotTableVersion15).CreatePivotChart(ChartDestination:="Sheet1").Select
```