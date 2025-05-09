# ListObject SharePointURL Property

## Business Description
Returns a String representing the URL of the SharePoint list for a given ListObject object. Read-only String.

## Behavior
Returns aStringrepresenting the URL of the SharePoint list for a givenListObjectobject. Read-onlyString.

## Example Usage
```vba
Sub PublishList() 
 Dim wrksht As Worksheet 
 Dim objListObj As ListObject 
 Dim arTarget(4) As String 
 Dim strSTSConnection As String 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListObj = wrksht.ListObjects(1) 
 
 arTarget(0) = "0" 
 arTarget(1) = objListObj.SharePointURLarTarget(2) = "1" 
 arTarget(3) = objListObj.Name 
 
 strSTSConnection = objListObj.Publish(arTarget, True) 
End Sub
```