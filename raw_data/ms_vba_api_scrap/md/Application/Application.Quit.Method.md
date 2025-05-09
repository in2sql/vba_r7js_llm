# Application.Quit method (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
For Each w In Application.Workbooks 
 w.Save 
Next w 
Application.Quit
```

## Remarks
If unsaved workbooks are open when you use this method, Excel displays a dialog box asking whether you want to save the changes. You can prevent this by saving all workbooks before using the Quit method or by setting the DisplayAlerts property to False. When this property is False, Excel doesn't display the dialog box when you quit with unsaved workbooks; it quits without saving them.

## Example
No VBA example available.
