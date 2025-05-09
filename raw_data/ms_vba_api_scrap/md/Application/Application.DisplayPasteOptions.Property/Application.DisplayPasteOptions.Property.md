# Application.DisplayPasteOptions property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub CheckDisplayFeature() 
 
 ' Check if the options button can be displayed. 
 If Application.DisplayPasteOptions = True Then 
 MsgBox "The ability to display the Paste Options button is on." 
 Else 
 MsgBox "The ability to display the Paste Options button is off." 
 End If 
 
End Sub
```

## Remarks
This is a Microsoft Office-wide setting. This setting affects all other Microsoft Office applications. Setting the DisplayPasteOptions property to True turns off the Auto Fill Options button in Microsoft Excel. The Auto Fill Options button is only in Excel, but the Paste Options button is in all the other Microsoft Office applications.

## Example
```vba
Sub CheckDisplayFeature() 
 
 ' Check if the options button can be displayed. 
 If Application.DisplayPasteOptions = True Then 
 MsgBox "The ability to display the Paste Options button is on." 
 Else 
 MsgBox "The ability to display the Paste Options button is off." 
 End If 
 
End Sub
```

