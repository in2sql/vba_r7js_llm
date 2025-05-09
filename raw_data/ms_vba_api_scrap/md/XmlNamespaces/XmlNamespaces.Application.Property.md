# XmlNamespaces Application Property

## Business Description
When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.

## Behavior
When used without an object qualifier, this property returns anApplicationobject that represents the Microsoft Excel application. When used with an object qualifier, this property returns anApplicationobject that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.

## Example Usage
```vba
Set myObject = ActiveWorkbook 
If myObject.Application.Value = "Microsoft Excel" Then 
 MsgBox "This is an Excel Application object." 
Else 
 MsgBox "This is not an Excel Application object." 
End If
```