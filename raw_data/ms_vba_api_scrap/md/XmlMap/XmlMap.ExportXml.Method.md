# XmlMap ExportXml Method

## Business Description
Exports the contents of cells mapped to the specified XmlMap object to a String variable.

## Behavior
Exports the contents of cells mapped   to the specifiedXmlMapobject to aStringvariable.

## Example Usage
```vba
Sub ExportToString() 
 Dim strContactData As String 
 
 ActiveWorkbook.XmlMaps("Contacts").ExportXmlData:=strContactData 
End Sub
```