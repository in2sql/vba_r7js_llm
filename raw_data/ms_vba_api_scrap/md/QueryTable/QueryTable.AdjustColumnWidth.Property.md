# QueryTable AdjustColumnWidth Property

## Business Description
True if the column widths are automatically adjusted for the best fit each time you refresh the specified query table. False if the column widths are not automatically adjusted with each refresh. The default value is True. Read/write Boolean.

## Behavior
Trueif the column widths are automatically adjusted for the best fit each time you refresh the specified query table.Falseif the column widths are not automatically adjusted with each refresh. The default value isTrue. Read/writeBoolean.

## Example Usage
```vba
With Workbooks(1).Worksheets(1).QueryTables _ 
 .Add(Connection:= varDBConnStr, _ 
 Destination:=Range("B1"), _ 
 Sql:="Select Price From CurrentStocks " & _ 
 "Where Symbol = 'MSFT'") 
 .AdjustColumnWidth= False 
 .Refresh 
End With
```