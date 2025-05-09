# Application.StatusBar property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
oldStatusBar = Application.DisplayStatusBar 
Application.DisplayStatusBar = True 
Application.StatusBar = "Please be patient..." 
Workbooks.Open filename:="LARGE.XLS" 
Application.StatusBar = False 
Application.DisplayStatusBar = oldStatusBar
```

## Remarks
This property returns False if Microsoft Excel has control of the status bar. To restore the default status bar text, set the property to False; this works even if the status bar is hidden.

## Example
No VBA example available.
