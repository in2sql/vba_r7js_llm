# PublishObject AutoRepublish Property

## Business Description
When a workbook is saved, Microsoft Excel determines if any item in the PublishObjects collection has the AutoRepublish property set to True and, if so, republishes it. The default value is False. Read/write Boolean.

## Behavior
When a workbook is saved, Microsoft Excel determines if any item in thePublishObjectscollection has theAutoRepublishproperty set toTrueand, if so, republishes it. The default value isFalse. Read/writeBoolean.

## Example Usage
```vba
Sub PublishToWeb() 
 
 With ActiveWorkbook.PublishObjects.Add( _ 
 SourceType:= xlSourceRange, _ 
 Filename:="C:\Work.htm", _ 
 Sheet:="Sheet1", _ 
 Source:="A1:D10", _ 
 HtmlType:=xlHtmlStatic, _ 
 DivID:="Book1.xls_130489") 
 .Publish 
 .AutoRepublish= True 
 End With 
 
End Sub
```