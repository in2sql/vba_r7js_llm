# PublishObject DivID Property

## Business Description
Returns the unique identifier used for identifying an HTML <DIV> tag on a Web page. The tag is associated with an item in a document that you have saved to a Web page.

## Behavior
Returns the unique identifier used for identifying an HTML <DIV> tag on a Web page. The tag is associated with an item in a document that you have saved to a Web page. An item can be an entire workbook, a worksheet, a selected print range, an AutoFilter range, a range of cells, a chart, a PivotTable report, or a query table. Read-onlyString.

## Example Usage
```vba
Set objPO = ActiveWorkbook.PublishObjects.Add( _ 
 SourceType:=xlSourceRange, _ 
 Filename:="\\Server1\Reports\q198.htm", _ 
 Sheet:="Sheet1", _ 
 Source:="C2:D6", _ 
 HtmlType:=xlHtmlStatic) 
objPO.Publish 
strTargetDivID = objPO.DivIDOpen "\\Server1\Reports\q198.htm" For Input As #1 
Open "\\Server1\Reports\newq1.htm" For Output As #2 
While Not EOF(1) 
 Line Input #1, strFileLine 
 If InStr(strFileLine, strTargetDivID) > 0 And _ 
 InStr(strFileLine, "<div") > 0 Then 
 Print #2, "<!--Saved item-->" 
 End If 
 Print #2, strFileLine 
Wend 
Close #2 
Close #1
```