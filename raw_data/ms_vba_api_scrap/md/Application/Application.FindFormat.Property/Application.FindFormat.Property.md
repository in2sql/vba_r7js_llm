# Application.FindFormat property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub UseFindFormat() 
 
 ' Establish search criteria. 
 With Application.FindFormat.Font 
 .Name = "Arial" 
 .FontStyle = "Regular" 
 .Size = 10 
 End With 
 
 ' Notify user. 
 With Application.FindFormat.Font 
 MsgBox .Name & "-" & .FontStyle & "-" & .Size & _ 
 " font is what the search criteria is set to." 
 End With 
 
End Sub
```

## Example
```vba
Sub UseFindFormat() 
 
 ' Establish search criteria. 
 With Application.FindFormat.Font 
 .Name = "Arial" 
 .FontStyle = "Regular" 
 .Size = 10 
 End With 
 
 ' Notify user. 
 With Application.FindFormat.Font 
 MsgBox .Name & "-" & .FontStyle & "-" & .Size & _ 
 " font is what the search criteria is set to." 
 End With 
 
End Sub
```

