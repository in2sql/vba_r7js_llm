# Application.ReplaceFormat property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub MakeBold() 
 
 ' Establish search criteria. 
 With Application.FindFormat.Font 
 .Name = "Arial" 
 .FontStyle = "Regular" 
 .Size = 10 
 End With 
 
 ' Establish replacement criteria. 
 With Application.ReplaceFormat.Font 
 .Name = "Arial" 
 .FontStyle = "Bold" 
 .Size = 8 
 End With 
 
 ' Notify user. 
 With Application.ReplaceFormat.Font 
 MsgBox .Name & "-" & .FontStyle & "-" & .Size & _ 
 " font is what the search criteria will replace cell formats with." 
 End With 
 
 ' Make the replacements on the worksheet. 
 Cells.Replace What:="", Replacement:="", _ 
 SearchFormat:=True, ReplaceFormat:=True 
 
End Sub
```

## Example
```vba
Sub MakeBold() 
 
 ' Establish search criteria. 
 With Application.FindFormat.Font 
 .Name = "Arial" 
 .FontStyle = "Regular" 
 .Size = 10 
 End With 
 
 ' Establish replacement criteria. 
 With Application.ReplaceFormat.Font 
 .Name = "Arial" 
 .FontStyle = "Bold" 
 .Size = 8 
 End With 
 
 ' Notify user. 
 With Application.ReplaceFormat.Font 
 MsgBox .Name & "-" & .FontStyle & "-" & .Size & _ 
 " font is what the search criteria will replace cell formats with." 
 End With 
 
 ' Make the replacements on the worksheet. 
 Cells.Replace What:="", Replacement:="", _ 
 SearchFormat:=True, ReplaceFormat:=True 
 
End Sub
```

