# OLEDBConnection UseLocalConnection Property

## Business Description
True if the LocalConnection property is used to specify the string that enables Microsoft Excel to connect to a data source. False if the connection string specified by the Connection property is used. Read/write Boolean.

## Behavior
Trueif theLocalConnectionproperty is used to specify the string that enables Microsoft Excel to connect to a data source.Falseif the connection string specified by theConnectionproperty is used. Read/writeBoolean.

## Example Usage
```vba
With ActiveWorkbook.PivotCaches(1) 
 .LocalConnection = _ 
 "OLEDB;Provider=MSOLAP;Data Source=C:\Data\DataCube.cub" 
 .UseLocalConnection = True 
End With
```