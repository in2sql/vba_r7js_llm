# Worksheet CustomProperties Property

## Business Description
Returns a CustomProperties object representing the identifier information associated with a worksheet.

## Behavior
Returns aCustomPropertiesobject representing the identifier information associated with a worksheet.

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