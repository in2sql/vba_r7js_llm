# PublishObject Publish Method

## Business Description
Saves an item or a collection of items in a document to a Web page.

## Behavior
Saves an item or a collection of items in a document to a Web page.

## Example Usage
```vba
With ActiveWorkbook.PublishObjects.Add(xlSourceRange, _ 
 "\\Server1\sharedfolder\stockreport.htm", "First Quarter", _ 
 "$D$5:$D$9", xlHtmlStatic, "Book2_25082", "") 
 .Publish(True) 
 .AutoRepublish = True 
End With
```