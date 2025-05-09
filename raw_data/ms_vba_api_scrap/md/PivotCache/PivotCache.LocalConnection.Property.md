# PivotCache LocalConnection Property

## Business Description
Returns or sets the connection string to an offline cube file. Read/write String.

## Behavior
Returns or sets the connection string to an offline cube file. Read/writeString.

## Example Usage
```vba
With ActiveWorkbook.PivotCaches(1) 
 .LocalConnection= _ 
 "OLEDB;Provider=MSOLAP;Data Source=C:\Data\DataCube.cub" 
 .UseLocalConnection = True 
End With
```