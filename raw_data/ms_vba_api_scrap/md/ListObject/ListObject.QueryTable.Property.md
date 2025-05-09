# ListObject QueryTable Property

## Business Description
Returns the QueryTable object that provides a link for the ListObject object to the list server. Read-only.

## Behavior
Returns theQueryTableobject  that provides a link for theListObjectobject to the list server. Read-only.

## Example Usage
```vba
Dim wrksht As Worksheet 
 Dim objListObj As ListObject 
 Dim objQryTbl As QueryTable 
 Dim prpQryProp As pro 
 Dim arTarget(4) As String 
 Dim strSTSConnection As String 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListObj = wrksht.ListObjects(1) 
 
 arTarget(0) = "0" 
 arTarget(1) = "http://myteam/project1" 
 arTarget(2) = "1" 
 arTarget(3) = "List1" 
 
 strSTSConnection = objListObj.Publish(arTarget, True) 
 
 Set objQryTbl = objListObj.QueryTable 
 
 objQryTbl.MaintainConnection = True
```