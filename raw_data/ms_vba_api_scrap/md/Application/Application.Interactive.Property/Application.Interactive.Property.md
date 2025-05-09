# Application.Interactive property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Application.Interactive = False 
Application.DisplayAlerts = False 
channelNumber = Application.DDEInitiate( _ 
 app:="WinWord", _ 
 topic:="C:\WINWORD\FORMLETR.DOC") 
Application.DDEExecute channelNumber, "[FILEPRINT]" 
Application.DDETerminate channelNumber 
Application.DisplayAlerts = True 
Application.Interactive = True
```

## Remarks
Blocking user input prevents the user from interfering with the macro as it moves or activates Excel objects.

## Example
No VBA example available.
