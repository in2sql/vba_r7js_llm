# Application.DisplayXMLSourcePane method (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Application.CommandBars("XML Source").Visible = False
```

## Parameters
- **XmlMap**: Optional

## Remarks
Use the following code to hide the XML Source task pane.

## Example
```vba
Sub DisplayXMLMap() 
 Dim objCustomer As XmlMap 
 
 Set objCustomer = ActiveWorkbook.XmlMaps.Add( _ 
 "Customers.xsd", "Root") 
 
 objCustomer.Name = "Customers" 
 
 Application.DisplayXMLSourcePane 
 objCustomer 
End Sub
```

