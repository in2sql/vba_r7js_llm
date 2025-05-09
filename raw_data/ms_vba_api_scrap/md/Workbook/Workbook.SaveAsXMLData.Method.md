# Workbook SaveAsXMLData Method

## Business Description
Exports the data that has been mapped to the specified XML schema map to an XML data file.

## Behavior
Exports the data that has been mapped to the specified XML schema map to an XML data file.

## Example Usage
```vba
Sub ExportAsXMLData() 
 Dim objMapToExport As XmlMap 
 
 Set objMapToExport = ActiveWorkbook.XmlMaps("Customer") 
 
 If objMapToExport.IsExportable Then 
 
 ActiveWorkbook.SaveAsXMLData"Customer Data.xml", objMapToExport 
 Else 
 MsgBox "Cannot use " & objMapToExport.Name & _ 
 "to export the contents of the worksheet to XML data." 
 End If 
End Sub
```