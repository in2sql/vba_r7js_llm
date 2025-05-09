# Workbook PublishObjects Property

## Business Description
Returns the PublishObjects collection. Read-only.

## Behavior
Returns thePublishObjectscollection. Read-only.

## Example Usage
```vba
Set objPObjs = ActiveWorkbook.PublishObjectsFor Each objPO in objPObjs 
 If objPO.HtmlType = xlHTMLStatic Then 
 objPO.Publish 
 End If 
Next objPO
```