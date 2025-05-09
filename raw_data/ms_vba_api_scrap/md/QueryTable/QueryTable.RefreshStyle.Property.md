# QueryTable RefreshStyle Property

## Business Description
Returns or sets the way rows on the specified worksheet are added or deleted to accommodate the number of rows in a recordset returned by a query. Read/write XlCellInsertionMode.

## Behavior
Returns or sets the way rows on the specified worksheet are added or deleted to accommodate the number of rows in a recordset returned by a query. Read/writeXlCellInsertionMode.

## Example Usage
```vba
Dim qt As QueryTable 
Set qt = Sheets("sheet1").QueryTables _ 
 .Add(Connection:="Finder;c:\myfile.dqy", _ 
 Destination:=Range("sheet1!a1")) 
With qt 
 .RefreshStyle= xlInsertEntireRows 
 .Refresh 
End With
```