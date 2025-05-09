# Application.ShowStartupDialog property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub CheckStartupDialog() 
 
 ' Determine if the New Workbook task pane is enabled. 
 If Application.ShowStartupDialog = False Then 
 MsgBox "ShowStartupDialog is set to False." 
 Else 
 MsgBox "ShowStartupDialog is set to True." 
 End If 
 
End Sub
```

## Example
```vba
Sub CheckStartupDialog() 
 
 ' Determine if the New Workbook task pane is enabled. 
 If Application.ShowStartupDialog = False Then 
 MsgBox "ShowStartupDialog is set to False." 
 Else 
 MsgBox "ShowStartupDialog is set to True." 
 End If 
 
End Sub
```

