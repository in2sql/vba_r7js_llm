# OLEDBConnection SaveAsODC Method

## Business Description
Saves the OLE DB connection as an Microsoft Office Data Connection file.

## Behavior
Saves the OLE DB connection as an Microsoft Office Data Connection file.

## Example Usage
```vba
Sub UseSaveAsODC() 
 
 Application.ActiveWorkbook.OLEDBConnection.SaveAsODC ("ODCFile") 
 
End Sub
```