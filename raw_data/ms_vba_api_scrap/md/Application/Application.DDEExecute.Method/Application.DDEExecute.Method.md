# Application.DDEExecute method (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
channelNumber = Application.DDEInitiate( _ 
 app:="WinWord", _ 
 topic:="C:\WINWORD\FORMLETR.DOC") 
Application.DDEExecute channelNumber, "[FILEPRINT]" 
Application.DDETerminate channelNumber
```

## Parameters
- **Channel**: Required
- **String**: Required

## Remarks
The DDEExecute method is designed to send commands to another application. You can also use it to send keystrokes to another application, although the SendKeys method is the preferred way to send keystrokes.

## Example
No VBA example available.
