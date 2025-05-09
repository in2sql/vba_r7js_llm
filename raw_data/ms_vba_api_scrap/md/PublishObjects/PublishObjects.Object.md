# PublishObjects Object

## Business Description
A collection of all PublishObject objects in the workbook.

## Behavior
A collection of allPublishObjectobjects in the workbook.

## Example Usage
```vba
Set objPObjs = ActiveWorkbook.PublishObjectsFor Each objPO in objPObjs 
 If objPO.HtmlType = xlHTMLStatic Then 
 objPO.Publish 
 End If 
Next objPO
```