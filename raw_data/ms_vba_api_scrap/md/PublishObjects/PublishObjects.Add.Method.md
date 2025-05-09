# PublishObjects Add Method

## Business Description
Creates an object that represents an item in a document saved to a Web page. Such objects facilitate subsequent updates to the Web page while automated changes are being made to the document in Microsoft Excel. Returns a PublishObject object.

## Behavior
Creates an object that represents an item in a document saved to a Web page. Such objects facilitate subsequent updates to the Web page while automated changes are being made to the document in Microsoft Excel. Returns aPublishObjectobject.

## Example Usage
```vba
With ActiveWorkbook.PublishObjects.Add(SourceType:=xlSourceRange, _ 
    Filename:="\\Server\Stockreport.htm", Sheet:="First Quarter", Source:="$G$3:$H$6", _ 
    HtmlType:=xlHtmlStatic, DivID:="Book1_4170") 
        .Publish (True) 
        .AutoRepublish = False 
End With
```