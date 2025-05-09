# TableObject Object

## Business Description
Represents a worksheet table built from data returned from a PowerPivot model.

## Behavior
Represents a worksheet table built from data returned from a PowerPivot model.

## Example Usage
```vba
Sub CreateTable()
Dim objWBConnection As WorkbookConnection
Dim objWorksheet As Worksheet
Dim objTable As TableObject   'This is the new Table object

Set objWorksheet = ActiveWorkbook.Worksheets("Sheet1")

'Create a WorkbookConnection to the external data source first.
Set objWBConnection = ActiveWorkbook.Connections.Add2( _
        "Cubes3 AdventureWorksDW DimEmployee1", "", Array( _
        "OLEDB;Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=AdventureWorksDW;Data Source=MyServer;Use " _
        , _
        "Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=MYWORKSTATION;Use Encryption for Data=False;Tag with co" _
        , "lumn collation when possible=False"), Array( _
        """AdventureWorksDW"".""dbo"".""DimEmployee"""), 3, True)

'Create a new table connected to the model.
Set objTable = objWorksheet.ListObjects.Add(SourceType:=xlSrcModel, Source:=objWBConnection, Destination:=Range("$A$1")).TableObject

objTable.Refresh

End Sub
```