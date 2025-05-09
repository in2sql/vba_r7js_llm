# PivotCache UseLocalConnection Property

## Business Description
Returns True if the LocalConnection property is used to specify the string that enables Microsoft Excel to connect to a data source. Returns False if the connection string specified by the Connection property is used. Read/write Boolean.

## Behavior
ReturnsTrueif theLocalConnectionproperty is used to specify the string that enables Microsoft Excel to connect to a data source. ReturnsFalseif the connection string specified by theConnectionproperty is used. Read/writeBoolean.

## Example Usage
```vba
With ActiveWorkbook.PivotCaches(1) 
 .LocalConnection = _ 
 "OLEDB;Provider=MSOLAP;Data Source=C:\Data\DataCube.cub" 
 .UseLocalConnection = True 
End With
```