# XPath SetValue Method

## Business Description
Maps the specified XPath object to a ListColumn object or Range collection. If the XPath object has previously been mapped to the ListColumn object or Range collection, the SetValue method sets the properties of the XPath object.

## Behavior
Maps the specifiedXPathobject to aListColumnobject orRangecollection. If theXPathobject has previously been mapped to theListColumnobject orRangecollection, theSetValuemethod sets the properties of theXPathobject.

## Example Usage
```vba
Sub CreateXMLList() 
    Dim mapContact As XmlMap 
    Dim strXPath As String 
    Dim lstContacts As ListObject 
    Dim objNewCol As ListColumn 
 
    ' Specify the schema map to use. 
    Set mapContact = ActiveWorkbook.XmlMaps("Contacts") 
     
    ' Create a new list. 
    Set lstContacts = ActiveSheet.ListObjects.Add 
         
    ' Specify the first element to map. 
    strXPath = "/Root/Person/FirstName" 
    ' Map the element. 
    lstContacts.ListColumns(1).XPath.SetValuemapContact, strXPath 
 
    ' Specify the second element to map. 
    strXPath = "/Root/Person/LastName" 
    ' Add a column to the list. 
    Set objNewCol = lstContacts.ListColumns.Add 
    ' Map the element. 
    objNewCol.XPath.SetValuemapContact, strXPath 
 
    strXPath = "/Root/Person/Address/Zip" 
    Set objNewCol = lstContacts.ListColumns.Add 
    objNewCol.XPath.SetValuemapContact, strXPath 
End Sub
```