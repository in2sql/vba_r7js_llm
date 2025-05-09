# Workbook Connections Property

## Business Description
The Connections property establishes a connection between the workbook and an ODBC or an OLEDB data source and refreshes the data without prompting the user. Read-only.

## Behavior
TheConnectionsproperty establishes a connection between the workbook and an ODBC or an OLEDB data source and refreshes the data without prompting the user. Read-only.

## Example Usage
```vba
ActiveWorkbook.Connections(1).ODBCConnection.Refresh 
ActiveWorkbook.Connections(1).OLEDBConnection.Refresh
```