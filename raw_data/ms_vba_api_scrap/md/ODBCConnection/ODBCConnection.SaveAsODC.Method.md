# ODBCConnection SaveAsODC Method

## Business Description
Saves the ODBC connection as a Microsoft Office Data Connection file.

## Behavior
Saves the ODBC connection as a Microsoft Office Data Connection file.

## Example Usage
```vba
Sub UseSaveAsODC() 
 
 Application.ActiveWorkbook.ODBCConnection.SaveAsODC ("ODCFile") 
 
End Sub
```