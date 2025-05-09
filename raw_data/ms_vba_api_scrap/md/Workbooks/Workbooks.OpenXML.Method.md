# Workbooks OpenXML Method

## Business Description
Opens an XML data file. Returns a Workbook object.

## Behavior
Opens an XML data file. Returns aWorkbookobject.

## Example Usage
```vba
Sub UseOpenXML() 
 Application.Workbooks.OpenXML_ 
 Filename:="customers.xml", _ 
 LoadOption:=xlXmlLoadImportToList 
End Sub
```