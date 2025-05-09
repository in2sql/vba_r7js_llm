# QueryTable PreserveColumnInfo Property

## Business Description
True if column sorting, filtering, and layout information is preserved whenever a query table is refreshed. The default value is True. Read/write Boolean.

## Behavior
Trueif column sorting, filtering, and layout information is preserved whenever a query table is refreshed. The default value isTrue. Read/writeBoolean.

## Example Usage
```vba
Dim cnnConnect As ADODB.Connection 
Dim rstRecordset As ADODB.Recordset 
 
Set cnnConnect = New ADODB.Connection 
cnnConnect.Open "Provider=SQLOLEDB;" & _ 
 "Data Source=srvdata;" & _ 
 "User ID=wadet;Password=4me2no;" 
 
Set rstRecordset = New ADODB.Recordset 
rstRecordset.Open _ 
 Source:="Select Name, Quantity, Price From Products", _ 
 ActiveConnection:=cnnConnect, _ 
 CursorType:=adOpenDynamic, _ 
 LockType:=adLockReadOnly, _ 
 Options:=adCmdText 
 
With ActiveSheet.QueryTables.Add( _ 
 Connection:=rstRecordset, _ 
 Destination:=Range("A1")) 
 .Name = "Contact List" 
 .FieldNames = True 
 .RowNumbers = False 
 .FillAdjacentFormulas = False 
 .PreserveFormatting = True 
 .RefreshOnFileOpen = False 
 .BackgroundQuery = True 
 .RefreshStyle = xlInsertDeleteCells 
 .SavePassword = True 
 .SaveData = True 
 .AdjustColumnWidth = True 
 .RefreshPeriod = 0 
 .PreserveColumnInfo= True 
 .Refresh BackgroundQuery:=False 
End With
```