# DataFeedConnection SaveAsODC Method

## Business Description
Saves the data feed connection as a Microsoft Office Data Connection file.

## Behavior
Saves the data feed connection as a Microsoft Office Data Connection file.

## Example Usage
```vba
Sub UseSaveAsODC() 
 
   Application.ActiveWorkbook.Connections("Datafeed1").DataFeedConnection.SaveAsODC ("ODCFile")
 
End Sub
```