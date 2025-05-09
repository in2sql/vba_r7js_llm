# XPath Object

## Business Description
Represents an XPath that has been mapped to a Range or ListColumn object.

## Behavior
Represents an XPath that has been mapped to aRangeorListColumnobject.

## Example Usage
```vba
Sub CreateXMLList() 
 Dim mapContact As XmlMap 
 Dim strXPath As String 
 Dim lstContacts As ListObject 
 Dim lcNewCol As ListColumn 
 
 ' Specify the schema map to use. 
 Set mapContact = ActiveWorkbook.XmlMaps("Contacts") 
 
 ' Create a new list. 
 Set lstContacts = ActiveSheet.ListObjects.Add 
 
 ' Specify the first element to map. 
 strXPath = "/Root/Person/FirstName" 
 ' Map the element. 
 lstContacts.ListColumns(1).XPath.SetValue mapContact, strXPath 
 
 ' Specify the element to map. 
 strXPath = "/Root/Person/LastName" 
 ' Add a column to the list. 
 Set lcNewCol = lstContacts.ListColumns.Add 
 ' Map the element. 
 lcNewCol.XPath.SetValue mapContact, strXPath 
 
 strXPath = "/Root/Person/Address/Zip" 
 Set lcNewCol = lstContacts.ListColumns.Add 
 lcNewCol.XPath.SetValue mapContact, strXPath 
End Sub
```