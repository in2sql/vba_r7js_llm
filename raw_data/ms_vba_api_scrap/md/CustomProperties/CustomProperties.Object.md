# CustomProperties Object

## Business Description
A collection of CustomProperty objects that represent additional information. The information can be used as metadata for XML.

## Behavior
A collection ofCustomPropertyobjects that represent additional information. The information can be used as metadata for XML.

## Example Usage
```vba
Sub CheckCustomProperties() 
 
 Dim wksSheet1 As Worksheet 
 
 Set wksSheet1 = Application.ActiveSheet 
 
 ' Add metadata to worksheet. 
 wksSheet1.CustomProperties.Add _ 
 Name:="Market", Value:="Nasdaq" 
 
 ' Display metadata. 
 With wksSheet1.CustomProperties.Item(1) 
 MsgBox .Name & vbTab & .Value 
 End With 
 
End Sub
```