# Workbooks OpenDatabase Method

## Business Description
Returns a Workbook object representing a database.

## Behavior
Returns aWorkbookobject representing a database.

## Example Usage
```vba
Sub UseOpenDatabase() 
 ' Open the Northwind database in the background and create a PivotTable 
 Workbooks.OpenDatabaseFilename:="c:\Northwind.mdb", _ 
 CommandText:="Orders", _ 
 CommandType:=xlCmdTable, _ 
 BackgroundQuery:=True, _ 
 ImportDataAs:=xlPivotTableReport 
End Sub
```