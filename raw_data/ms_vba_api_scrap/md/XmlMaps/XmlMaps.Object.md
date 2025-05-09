# XmlMaps Object

## Business Description
Represents the collection of XmlMap objects that have been added to a workbook.

## Behavior
Represents the collection ofXmlMapobjects that have been added to a workbook.

## Example Usage
```vba
Sub AddXmlMap() 
 Dim strSchemaLocation As String 
 
 strSchemaLocation = "http://example.microsoft.com/schemas/CustomerData.xsd" 
 ActiveWorkbook.XmlMaps.Add strSchemaLocation, "Root" 
End Sub
```